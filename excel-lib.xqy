
xquery version "1.0-ml";

module namespace excel-lib = "http://marklogic.com/solutions/htx/excel-lib";

import module namespace ooxml= "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy";
import module namespace functx = "http://www.functx.com" at "/MarkLogic/functx/functx-1.0-nodoc-2007-01.xqy";

declare namespace rels = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace r = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"; 

declare function excel-lib:load-report-template-file(
  $file,
  $collections)
{
  let $binary := xdmp:external-binary($file)
  let $_ := xdmp:document-insert($file,
    $binary,
    (),
    $collections)

  let $uris := ooxml:package-uris($binary)
  let $package-parts := ooxml:package-parts($binary)
    let $template-parts := 
      for $part at $index in $package-parts
      let $uri := $uris[$index]
      return
        xdmp:document-insert(
          fn:replace(fn:replace($file,"\.xlsx$","/")||$uri, "[/]{2,}","/"),
          if (fn:ends-with($uri, "/workbook.xml")) then
            excel-lib:add-calc-onload($part)
          else
            $part,
          (),
          $collections)

    return ()
};

declare function excel-lib:update-row(
  $rows as node()*,
  $details-map as map:map
)
{
  let $row := fn:head($rows)
  return
    if ($row instance of element(sheet:row)) then
      let $cell-ids := $row/sheet:c/@r
      let $range := 1 to fn:count($cell-ids)
      return (
        if (some $cid in $cell-ids satisfies map:contains($details-map, fn:string($cid))) then
          <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
            $row/@*,
            for $cell in $row/sheet:c
            let $cell-id := fn:string($cell/@r)
            let $value := (map:get($details-map, $cell-id), $cell/sheet:v)[1]/node()
            return 
              <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
                $cell/@* except $cell/@t,
                if (map:contains($details-map, $cell-id)) then
                  attribute t { 
                    if ($value castable as xs:integer or $value castable as xs:decimal 
                       or $value castable as xs:double or $value castable as xs:float) then
                      "n"
                    else
                      "str"
                  }
                else
                  $cell/@t,
                $cell/node() except $cell/sheet:v,
                <v xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
                  $value,
                  map:delete($details-map, $cell-id)
                }</v>
              }</c>
          }</row>
        else
          $row,
       excel-lib:update-row(
         fn:tail($rows),
         $details-map
       )
     )
    else if (map:count($details-map) eq 0) then (
      $rows
    )
    else if (fn:exists($row)) then (
      $row,
      excel-lib:update-row(
        fn:tail($rows),
        $details-map
      )
    ) else if (map:count($details-map) gt 0) then (
      let $row-map := map:map()
      let $_populate-rows :=
        for $key in map:keys($details-map)
        let $row-key := fn:replace($key, "^[A-Z]+([0-9]+)$", "$1")
        let $value := map:get($details-map, $key)/node()
        let $type := 
          if ($value castable as xs:integer or $value castable as xs:decimal 
             or $value castable as xs:double or $value castable as xs:float) then
            "n"
          else
            "str"
        return
          map:put(
            $row-map,
            $row-key,
            (
              map:get($row-map,$row-key),
              <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                r="{$key}" t="{$type}">
                <v xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
                  $value,
                  map:delete($details-map, $key)
                }</v>
              </c>
            )
          )
      for $row-key in map:keys($row-map)
      let $columns := map:get($row-map, $row-key)
      order by fn:number($row-key)
      return 
        <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
          attribute r { $row-key },
          for $col in $columns
          order by $col/@r
          return $col
        }</row>
    ) else ()
};


declare function excel-lib:update-sheet-with-results(
  $results,
  $sheet)
{
  let $details-map := map:new((
    for $detail in $results//detail
    return map:entry(fn:string($detail/@cell), $detail)
  ))
  let $sheet := $sheet/(self::sheet:worksheet|sheet:worksheet)[1]
  let $sheet-data := $sheet/sheet:sheetData
  let $new-sheet :=
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{
      $sheet/@*,
      $sheet/namespace::node(),
      $sheet/node()[. << $sheet-data],
      element sheet:sheetData {
        for $r in excel-lib:update-row($sheet-data/*, $details-map)
        order by fn:number($r/@r)
        return $r
      },
      $sheet/node()[. >> $sheet-data]
    }</worksheet>
  let $_replace :=
    xdmp:node-replace($sheet, $new-sheet)
  return $new-sheet
};


declare function excel-lib:process-results-for-sheet(
  $results-uri,
  $sheet-uri,
  $template-uri,
  $matching-pattern)
{  
  let $res := fn:doc($results-uri)
  let $template := (fn:doc(fn:replace($template-uri, "\.xlsx$", "/") || $sheet-uri),fn:doc($sheet-uri))[1]
  let $report-template := fn:doc($template-uri)  
  let $uris := ooxml:package-uris($report-template)
  let $package-parts := ooxml:package-parts($report-template)
  let $all-parts :=
    for $part at $index in $package-parts
    let $uri := $uris[$index]
    return 
      if (fn:ends-with($uri, $matching-pattern) ) then
        excel-lib:update-sheet-with-results($res, $template)
      else
        $part
    
  (: return $all-parts :)
  return ()
};

declare function excel-lib:generate-report(
  $template-uri,
  $output-report-uri,
  $collections)
{  
  let $report-template := fn:doc($template-uri)
  let $uris := ooxml:package-uris($report-template)
  let $package-parts := ooxml:package-parts($report-template)
  
  let $parts :=
    for $uri in $uris
    return (fn:doc(fn:replace($template-uri, "\.xlsx$", "/") || $uri),fn:doc($uri))[1]
      
  let $manifest :=
    <parts xmlns="xdmp:zip">
    {
      for $uri in $uris
      return <part>{$uri}</part>
    }
    </parts>
  
  let $_log := xdmp:log($manifest)
  let $pkg := xdmp:zip-create($manifest, ($parts))
  let $_ := xdmp:document-insert($output-report-uri, $pkg, (), $collections)
    
  return ()
};

declare function excel-lib:add-calc-onload($workbook) 
{
  let $workbook := ($workbook/(self::sheet:workbook|sheet:workbook))[1]
  let $definedNames := $workbook/sheet:definedNames
  let $calcPr := $workbook/sheet:calcPr
  let $newCalcPr := 
      element sheet:calcPr {
        attribute calcId { "122211" },
        attribute fullCalcOnLoad { 1 }
      }
  return
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >{
      $workbook/@*,
      $workbook/namespace::node(),
      if ($calcPr) then  (
        $workbook/node()[. << $calcPr],
        $newCalcPr,
        $workbook/node()[. >> $calcPr]
      ) else (
        $workbook/node()[. << $definedNames],
        $definedNames,
        $newCalcPr,
        $workbook/node()[. >> $definedNames]
      )
    }</workbook>
};


  