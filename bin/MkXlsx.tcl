# -*- tcl -*-
# mkxlsx.tcl -- Copyright (C) 2009 Pat Thoyts <patthoyts@users.sourceforge.net>
#               Copyright (C) 2024 Roman Dmytrenko
#
#        Create XLSX in Tcl.
# Create simple xlsx file from unpacked content
# Use mkzip package for pack xlsx zip file.
#
# version 1.0.0

package require Tcl 8.6

namespace eval ::zipfile::decode {}
namespace eval ::zipfile::encode {}
namespace eval ::mkxlsx {}

# 56,70 = 1mm
set ::mkxlsx::ooOneMm 56.70
set ::mkxlsx::ooLang "uk-UA"
set ::mkxlsx::ooFontOnePt 2
set ::mkxlsx::ooOneSizeBetween 240
set ::mkxlsx::ooGlueSpace 0x0A0

proc ::mkxlsx::setbinary chan {
  fconfigure $chan \
      -encoding    binary \
      -translation binary \
      -eofchar     {}

}

proc ::mkxlsx::timet_to_dos {time_t} {
    set s [clock format $time_t -format {%Y %m %e %k %M %S}]
    scan $s {%d %d %d %d %d %d} year month day hour min sec
    expr {(($year-1980) << 25) | ($month << 21) | ($day << 16)
          | ($hour << 11) | ($min << 5) | ($sec >> 1)}
}

proc ::mkxlsx::pop {varname {nth 0}} {
    upvar $varname args
    set r [lindex $args $nth]
    set args [lreplace $args $nth $nth]
    return $r
}

proc ::mkxlsx::walk {base {excludes ""} {match *} {path {}}} {
    set result {}
    set imatch [file join $path $match]
    set files [glob -nocomplain -tails -types f -directory $base -- $imatch]
    foreach file $files {
        set excluded 0
        foreach glob $excludes {
            if {[string match $glob $file]} {
                set excluded 1
                break
            }
        }
        if {!$excluded} {lappend result $file}
    }
    foreach dir [glob -nocomplain -tails -types d -directory $base -- $imatch] {
        set subdir [walk $base $excludes $match $dir]
        if {[llength $subdir]>0} {
            set result [concat $result [list $dir] $subdir]
        }
    }
    return $result
}

proc ::mkxlsx::add_file_to_archive {zipchan base path {comment ""}} {
    set fullpath [file join $base $path]
    set mtime [timet_to_dos [file mtime $fullpath]]
    if {[file isdirectory $fullpath]} {
        append path /
    }
    set utfpath [encoding convertto utf-8 $path]
    set utfcomment [encoding convertto utf-8 $comment]
    set flags [expr {(1<<11)}] ;# utf-8 comment and path
    set method 0               ;# store 0, deflate 8
    set attr 0                 ;# text or binary (default binary)
    set version 20             ;# minumum version req'd to extract
    set extra ""
    set crc 0
    set size 0
    set csize 0
    set data ""
    set seekable [expr {[tell $zipchan] != -1}]
    if {[file isdirectory $fullpath]} {
        set attrex 0x41ff0010  ;# 0o040777 (drwxrwxrwx)
    } elseif {[file executable $fullpath]} {
        set attrex 0x81ff0080  ;# 0o100777 (-rwxrwxrwx)
    } else {
        set attrex 0x81b60020  ;# 0o100666 (-rw-rw-rw-)
        if {[file extension $fullpath] in {".tcl" ".txt" ".c"}} {
            set attr 1         ;# text
        }
    }

    if {[file isfile $fullpath]} {
        set size [file size $fullpath]
        if {!$seekable} {set flags [expr {$flags | (1 << 3)}]}
    }

    set offset [tell $zipchan]
    set local [binary format a4sssiiiiss PK\03\04 \
                   $version $flags $method $mtime $crc $csize $size \
                   [string length $utfpath] [string length $extra]]
    append local $utfpath $extra
    puts -nonewline $zipchan $local

    if {[file isfile $fullpath]} {
        # If the file is under 2MB then zip in one chunk, otherwize we use
        # streaming to avoid requiring excess memory. This helps to prevent
        # storing re-compressed data that may be larger than the source when
        # handling PNG or JPEG or nested ZIP files.
        if {$size < 0x00200000} {
            set fin [::open $fullpath rb]
            setbinary $fin
            set data [::read $fin]
            set crc [::zlib crc32 $data]
            set cdata [::zlib deflate $data]
            if {[string length $cdata] < $size} {
                set method 8
                set data $cdata
            }
            close $fin
            set csize [string length $data]
            puts -nonewline $zipchan $data
        } else {
            set method 8
            set fin [::open $fullpath rb]
            setbinary $fin
            set zlib [::zlib stream deflate]
            while {![eof $fin]} {
                set data [read $fin 4096]
                set crc [zlib crc32 $data $crc]
                $zlib put $data
                if {[string length [set zdata [$zlib get]]]} {
                    incr csize [string length $zdata]
                    puts -nonewline $zipchan $zdata
                }
            }
            close $fin
            $zlib finalize
            set zdata [$zlib get]
            incr csize [string length $zdata]
            puts -nonewline $zipchan $zdata
            $zlib close
        }

        if {$seekable} {
            # update the header if the output is seekable
            set local [binary format a4sssiiii PK\03\04 \
                           $version $flags $method $mtime $crc $csize $size]
            set current [tell $zipchan]
            seek $zipchan $offset
            puts -nonewline $zipchan $local
            seek $zipchan $current
        } else {
            # Write a data descriptor record
            set ddesc [binary format a4iii PK\7\8 $crc $csize $size]
            puts -nonewline $zipchan $ddesc
        }
    }

    set hdr [binary format a4ssssiiiisssssii PK\01\02 0x0317 \
                 $version $flags $method $mtime $crc $csize $size \
                 [string length $utfpath] [string length $extra]\
                 [string length $utfcomment] 0 $attr $attrex $offset]
    append hdr $utfpath $extra $utfcomment
    return $hdr
}

proc ::mkxlsx::mkzip {filename args} {
  array set opts {
      -zipkit 0 -runtime "" -comment "" -directory ""
      -exclude {}
      -verbose 0
  }

  while {[string match -* [set option [lindex $args 0]]]} {
      switch -exact -- $option {
          -verbose { set opts(-verbose) 1}
          -zipkit  { set opts(-zipkit) 1 }
          -comment { set opts(-comment) [encoding convertto utf-8 [pop args 1]] }
          -runtime { set opts(-runtime) [pop args 1] }
          -directory {set opts(-directory) [file normalize [pop args 1]] }
          -exclude {set opts(-exclude) [pop args 1] }
          -- { pop args ; break }
          default {
              break
          }
      }
      pop args
  }

  set zf [::open $filename wb]
  setbinary $zf
  if {$opts(-runtime) ne ""} {
      set rt [::open $opts(-runtime) rb]
      setbinary $rt
      fcopy $rt $zf
      close $rt
  } elseif {$opts(-zipkit)} {
      set zkd "#!/usr/bin/env tclkit\n\# This is a zip-based Tcl Module\n"
      append zkd "package require vfs::zip\n"
      append zkd "vfs::zip::Mount \[info script\] \[info script\]\n"
      append zkd "if {\[file exists \[file join \[info script\] main.tcl\]\]} \{\n"
      append zkd "    source \[file join \[info script\] main.tcl\]\n"
      append zkd "\}\n"
      append zkd \x1A
      puts -nonewline $zf $zkd
  }

  set count 0
  set cd ""

  if {$opts(-directory) ne ""} {
      set paths [walk $opts(-directory) $opts(-exclude)]
  } else {
      set paths [glob -nocomplain {*}$args]
  }

  if {[lindex $paths 1] == "docProps"} {
      lappend paths "_rels/.rels"
  }

  foreach path $paths {
      if {[string is true $opts(-verbose)]} {
        puts $path
      }
      append cd [add_file_to_archive $zf $opts(-directory) $path]
      incr count
  }
  set cdoffset [tell $zf]
  set endrec [binary format a4ssssiis PK\05\06 0 0 \
                  $count $count [string length $cd] $cdoffset\
                  [string length $opts(-comment)]]
  append endrec $opts(-comment)
  puts -nonewline $zf $cd
  puts -nonewline $zf $endrec
  close $zf

  return
}

proc ::mkxlsx::xlsxInit { {path ""} } \
{
    if {$path == ""} {
        set path $::workingDir/etc/xml_templates
    }
    set ::mkxlsx::sharedStringsDoc [dom parse [tDOM::xmlReadFile [file nativename [file join $path sharedStrings.xml]]]]
    set ::mkxlsx::worksheetDoc [dom parse [tDOM::xmlReadFile [file nativename [file join $path sheet.xml]]]]
    set ::mkxlsx::styleSheetDoc [dom parse [tDOM::xmlReadFile [file nativename [file join $path styles.xml]]]]
    set ::mkxlsx::workbookDoc [dom parse [tDOM::xmlReadFile [file nativename [file join $path workbook.xml]]]]
    set ::mkxlsx::sharedStringsRoot [$::mkxlsx::sharedStringsDoc documentElement]
    set ::mkxlsx::worksheetRoot     [$::mkxlsx::worksheetDoc documentElement]
    set ::mkxlsx::styleSheetRoot    [$::mkxlsx::styleSheetDoc documentElement]
    set ::mkxlsx::workbookRoot      [$::mkxlsx::workbookDoc documentElement]
    set ::mkxlsx::defaultHeight 18.75
    set ::mkxlsx::defaultWidth 11.53
    set ::mkxlsx::boldStyle 3
    set ::mkxlsx::headerStyle 4
    set ::mkxlsx::headerFilledStyle 11
    set ::mkxlsx::textStyle 10
    set ::mkxlsx::dateStyle 9
    set ::mkxlsx::emptyStyle 2
    set ::mkxlsx::maxCol "A"
    set ::mkxlsx::worksheet [dict create]
    set ::mkxlsx::rows [dict create]
    set ::mkxlsx::cols [dict create]
    set ::mkxlsx::styles [dict create]
    set ::mkxlsx::fonts [dict create]
    set ::mkxlsx::borders [dict create]
    set ::mkxlsx::fills [dict create]
 
    set ::mkxlsx::mergeCells [$::mkxlsx::worksheetRoot getElementsByTagName mergeCells]
    set ::mkxlsx::colsNodes [$::mkxlsx::worksheetRoot getElementsByTagName sheetData]
    set ::mkxlsx::sheetView [$::mkxlsx::worksheetRoot getElementsByTagName sheetView]
}

proc ::mkxlsx::addRow { row {style 2} {ht 18.75} {customFormat 1} {hidden 0} {customHeight 1} {outlineLevel 0} {collapsed 0} } \
{
    dict set ::mkxlsx::rows $row style $style
    dict set ::mkxlsx::rows $row customFormat $customFormat
    dict set ::mkxlsx::rows $row ht $ht
    dict set ::mkxlsx::rows $row hidden $hidden
    dict set ::mkxlsx::rows $row customHeight $customHeight
    dict set ::mkxlsx::rows $row outlineLevel $outlineLevel
    dict set ::mkxlsx::rows $row collapsed $collapsed
}

proc ::mkxlsx::addCell { col row value {style 2} {type "s"} {formula ""}} \
{
    if {![dict exists $::mkxlsx::rows $row]} {
        ::mkxlsx::addRow $row $style
    }

    if {$type == "s" && $value != ""} {
        $::mkxlsx::sharedStringsRoot appendXML "<si><t xml:space=\"preserve\">$value</t></si>"
        dict set ::mkxlsx::worksheet $row $col value [expr {[llength [$::mkxlsx::sharedStringsRoot childNodes]] - 1}] 
        dict set ::mkxlsx::worksheet $row $col type $type 
    } else {
        dict set ::mkxlsx::worksheet $row $col value $value
    }

    dict set ::mkxlsx::worksheet $row $col style $style 
    dict set ::mkxlsx::worksheet $row $col formula $formula 

    if {[string toupper $col] > $::mkxlsx::maxCol} {
        set ::mkxlsx::maxCol [string toupper $col]
    }
}

proc ::mkxlsx::merge { dime } \
{
    $::mkxlsx::mergeCells appendXML "<mergeCell ref=\"$dime\"/>"
}

proc ::mkxlsx::setRowHeight {row {height 18.75}} \
{
    if {![dict exists $::mkxlsx::rows $row style]} {
        ::mkxlsx::addRow $row         
    }
    dict set ::mkxlsx::rows $row ht $height
}

proc ::mkxlsx::setColWidth {col {width 11.53}} \
{
    dict set ::mkxlsx::cols $col width $width
}

proc ::mkxlsx::createFont {args} \
{
    set index [expr {[dict size $::mkxlsx::fonts] + 1}]
    dict set ::mkxlsx::fonts $index sz "14" 
    dict set ::mkxlsx::fonts $index name "Times New Roman" 
    dict set ::mkxlsx::fonts $index family "0" 
    dict set ::mkxlsx::fonts $index charset "1"
    foreach {key value} $args {
        dict set ::mkxlsx::fonts $index $key $value
    }
    return [expr {[[$::mkxlsx::styleSheetRoot getElementsByTagName "fonts"] getAttribute "count"] + [dict size $::mkxlsx::fonts] - 1}]
}

proc ::mkxlsx::createFill {args} \
{
    set index [expr {[dict size $::mkxlsx::fills] + 1}]
    dict set ::mkxlsx::fills $index patternType "gray125"
    foreach {key value} $args {
        dict set ::mkxlsx::fills $index $key $value
    }
    return [expr {[[$::mkxlsx::styleSheetRoot getElementsByTagName "fills"] getAttribute "count"] + [dict size $::mkxlsx::fills] - 1}]
}

proc ::mkxlsx::splitSheet {{x 1} {y 1} {name A}} \
{
    $::mkxlsx::worksheetDoc createElement pane nPane
    $nPane setAttribute xSplit $x ySplit $y topLeftCell $name[expr {$y + 1}] activePane bottomRight state frozen
    $::mkxlsx::sheetView appendChild $nPane
    $::mkxlsx::worksheetDoc createElement selection nSel
    $nSel setAttribute pane topRight
    $::mkxlsx::sheetView appendChild $nSel
    $::mkxlsx::worksheetDoc createElement selection nSel
    $nSel setAttribute pane bottomLeft
    $::mkxlsx::sheetView appendChild $nSel
    $::mkxlsx::worksheetDoc createElement selection nSel
    $nSel setAttribute pane bottomRight activeCell $name[expr {$y + 1}] sqref $name[expr {$y + 1}]
    $::mkxlsx::sheetView appendChild $nSel
}

proc ::mkxlsx::createBorder {args} \
{
    set index [expr {[dict size $::mkxlsx::borders] + 1}]
    dict set ::mkxlsx::borders $index left "" 
    dict set ::mkxlsx::borders $index right "" 
    dict set ::mkxlsx::borders $index top "" 
    dict set ::mkxlsx::borders $index bottom "" 
    dict set ::mkxlsx::borders $index diagonal ""
    foreach {key value} $args {
        dict set ::mkxlsx::borders $index $key $value
    }
    return [expr {[[$::mkxlsx::styleSheetRoot getElementsByTagName "borders"] getAttribute "count"] + [dict size $::mkxlsx::borders] - 1}]
}

proc ::mkxlsx::createStyle {args} \
{
    set index [expr {[dict size $::mkxlsx::styles] + 1}]
    dict set ::mkxlsx::styles $index numFmtId "164" 
    dict set ::mkxlsx::styles $index fontId "5" 
    dict set ::mkxlsx::styles $index fillId "0" 
    dict set ::mkxlsx::styles $index borderId "0" 
    dict set ::mkxlsx::styles $index xfId "0" 
    dict set ::mkxlsx::styles $index applyFont "1" 
    dict set ::mkxlsx::styles $index applyBorder "1" 
    dict set ::mkxlsx::styles $index applyAlignment "1" 
    dict set ::mkxlsx::styles $index applyProtection "1" 
    dict set ::mkxlsx::styles $index horizontal "general" 
    dict set ::mkxlsx::styles $index vertical "bottom" 
    dict set ::mkxlsx::styles $index textRotation "0" 
    dict set ::mkxlsx::styles $index wrapText "0" 
    dict set ::mkxlsx::styles $index ident "0" 
    dict set ::mkxlsx::styles $index shrinkToFit "0" 
    dict set ::mkxlsx::styles $index locked "1" 
    dict set ::mkxlsx::styles $index hidden "0"
    foreach {key value} $args {
        dict set ::mkxlsx::styles $index $key $value
    }
    return [expr {[[$::mkxlsx::styleSheetRoot getElementsByTagName "cellXfs"] getAttribute "count"] + [dict size $::mkxlsx::styles] - 1}]
}

proc ::mkxlsx::setTabName {name} \
{
    set wbSheet [$::mkxlsx::workbookRoot getElementsByTagName "sheet"]
    $wbSheet setAttribute name $name
}

proc ::mkxlsx::writeXml {} \
{
    $::mkxlsx::sharedStringsRoot setAttribute count [llength [$::mkxlsx::sharedStringsRoot childNodes]]
    $::mkxlsx::sharedStringsRoot setAttribute uniqueCount [llength [$::mkxlsx::sharedStringsRoot childNodes]]
    $::mkxlsx::mergeCells setAttribute count [llength [$::mkxlsx::mergeCells childNodes]]

    set rows [lsort [dict keys $::mkxlsx::rows]]
    foreach key $rows {
        $::mkxlsx::worksheetDoc createElement row newRow
        $::mkxlsx::colsNodes appendChild $newRow
        $newRow setAttribute r $key s [dict get $::mkxlsx::rows $key style] ht [dict get $::mkxlsx::rows $key ht] \
                customFormat [dict get $::mkxlsx::rows $key customFormat] hidden [dict get $::mkxlsx::rows $key hidden] \
                customHeight [dict get $::mkxlsx::rows $key customHeight] outlineLevel [dict get $::mkxlsx::rows $key outlineLevel] \
                collapsed [dict get $::mkxlsx::rows $key collapsed]
        if {[dict exists $::mkxlsx::worksheet $key]} {
            set cols [lsort [dict keys [dict get $::mkxlsx::worksheet $key]]]
            foreach col $cols {
                $::mkxlsx::worksheetDoc createElement c newCol
                $::mkxlsx::worksheetDoc createElement v newValue
                $newCol setAttribute r $col$key s [dict get $::mkxlsx::worksheet $key $col style] 
            
                if {[dict exists $::mkxlsx::worksheet $key $col type]} {
                    $newCol setAttribute t [dict get $::mkxlsx::worksheet $key $col type]
                }
            
                if {[dict get $::mkxlsx::worksheet $key $col formula] != ""} {
                    $::mkxlsx::worksheetDoc createElement f newFormula
                    $newFormula appendChild [$::mkxlsx::worksheetDoc createTextNode [dict get $::mkxlsx::worksheet $key $col formula]]
                    $newCol appendChild $newFormula
                    $newValue appendChild [$::mkxlsx::worksheetDoc createTextNode ""]
                    $newCol appendChild $newValue
                } else {
                    if {[dict get $::mkxlsx::worksheet $key $col value] != ""} {
                        $newValue appendChild [$::mkxlsx::worksheetDoc createTextNode [dict get $::mkxlsx::worksheet $key $col value]]
                        $newCol appendChild $newValue
                    }
                }
                $newRow appendChild $newCol
            }
        }
    }

    set colsNode [$::mkxlsx::worksheetRoot getElementsByTagName cols]
    foreach col [$colsNode childNodes] {
        $colsNode removeChild $col
    }

    foreach key [dict keys $::mkxlsx::cols] {
        $::mkxlsx::worksheetDoc createElement col col
        $col setAttribute customWidth true style 2 max $key min $key width [dict get $::mkxlsx::cols $key width]
        $colsNode appendChild $col
        $col removeAttribute "xmlns"
    }

    if {[dict size $::mkxlsx::styles] > 0} {
        set  fontsNode [$::mkxlsx::styleSheetRoot getElementsByTagName fonts]
        foreach key [dict keys $::mkxlsx::fonts] {
            $::mkxlsx::styleSheetDoc createElement font font
            foreach {attr value} [dict get $::mkxlsx::fonts $key] {
                $::mkxlsx::styleSheetDoc createElement $attr node
                $node setAttribute val $value
                $font appendChild $node
            }
            $fontsNode appendChild $font
            $font removeAttribute "xmlns"
        }

        set fillsNode [$::mkxlsx::styleSheetRoot getElementsByTagName fills]
        foreach key [dict keys $::mkxlsx::fills] {
            $::mkxlsx::styleSheetDoc createElement fill fill
            $::mkxlsx::styleSheetDoc createElement patternFill pfill
            $pfill setAttribute patternType [dict get $::mkxlsx::fills $key patternType]
            if {[dict exists $::mkxlsx::fills $key fgColor]} {
                $::mkxlsx::styleSheetDoc createElement fgColor fgColor
                $fgColor setAttribute rgb [dict get $::mkxlsx::fills $key fgColor]
                $pfill appendChild $fgColor
            }
            if {[dict exists $::mkxlsx::fills $key bgColor]} {
                $::mkxlsx::styleSheetDoc createElement bgColor bgColor
                $bgColor setAttribute rgb [dict get $::mkxlsx::fills $key bgColor]
                $pfill appendChild $bgColor
            }
            $fill appendChild $pfill
            $fillsNode appendChild $fill
            $fill removeAttribute "xmlns"
        }

        set bordersNode [$::mkxlsx::styleSheetRoot getElementsByTagName borders]
        foreach key [dict keys $::mkxlsx::borders] {
            $::mkxlsx::styleSheetDoc createElement border border
            #$border setAttribute diagonalUp "false" diagonalDown "false"
            foreach {attr value} [dict get $::mkxlsx::borders $key] {
                $::mkxlsx::styleSheetDoc createElement $attr node
                $node setAttribute "style" $value
                $border appendChild $node
            }
            $bordersNode appendChild $border
            $border removeAttribute "xmlns"
        }

        set stylesNode [$::mkxlsx::styleSheetRoot getElementsByTagName cellXfs]
        foreach key [dict keys $::mkxlsx::styles] {
            $::mkxlsx::styleSheetDoc createElement xf xf
            $::mkxlsx::styleSheetDoc createElement alignment alignment
            $::mkxlsx::styleSheetDoc createElement protection protection

            $xf setAttribute numFmtId [dict get $::mkxlsx::styles $key numFmtId] fontId [dict get $::mkxlsx::styles $key fontId] fillId [dict get $::mkxlsx::styles $key fillId] borderId [dict get $::mkxlsx::styles $key borderId]\
                             xfId [dict get $::mkxlsx::styles $key xfId] applyFont [dict get $::mkxlsx::styles $key applyFont] applyBorder [dict get $::mkxlsx::styles $key applyBorder] \
                             applyAlignment [dict get $::mkxlsx::styles $key applyAlignment] applyProtection [dict get $::mkxlsx::styles $key applyProtection] 
            $alignment setAttribute horizontal [dict get $::mkxlsx::styles $key horizontal] vertical [dict get $::mkxlsx::styles $key vertical] textRotation [dict get $::mkxlsx::styles $key textRotation] \
                                    wrapText [dict get $::mkxlsx::styles $key wrapText] ident [dict get $::mkxlsx::styles $key ident] shrinkToFit [dict get $::mkxlsx::styles $key shrinkToFit] 
            $protection setAttribute locked [dict get $::mkxlsx::styles $key locked] hidden [dict get $::mkxlsx::styles $key hidden] 

            $xf appendChild $alignment
            $xf appendChild $protection

            $stylesNode appendChild $xf
            $xf removeAttribute "xmlns"
        }

        $fontsNode setAttribute count [llength [$fontsNode childNodes]]
        $fillsNode setAttribute count [llength [$fillsNode childNodes]]
        $bordersNode setAttribute count [llength [$bordersNode childNodes]]
        $stylesNode setAttribute count [llength [$stylesNode childNodes]]

        set fp [open [file nativename [file join $::workingDir etc xlsx xlsx_data xl styles.xml]]  w+]
        puts $fp [$::mkxlsx::styleSheetDoc asXML -xmlDeclaration 1]
        close $fp 
    }

    set dimen [$::mkxlsx::worksheetRoot getElementsByTagName dimension]
    $dimen setAttribute ref "A1:$::mkxlsx::maxCol[llength [$::mkxlsx::colsNodes childNodes]]"

    foreach object [list $::mkxlsx::sharedStringsRoot $::mkxlsx::colsNodes $::mkxlsx::mergeCells $::mkxlsx::sheetView] {
        if ([info exists object]) {
            foreach node [$object childNodes] {
                 $node removeAttribute "xmlns"            
            }
        }
    }

    set fp [open [file nativename [file join $::workingDir etc xlsx xlsx_data xl workbook.xml]]  w+]
    puts $fp [$::mkxlsx::workbookDoc asXML -xmlDeclaration 1]
    close $fp 

    set fp [open [file nativename [file join $::workingDir etc xlsx xlsx_data xl sharedStrings.xml]]  w+]
    puts $fp [$::mkxlsx::sharedStringsDoc asXML -xmlDeclaration 1]
    close $fp 
        
    set fp [open [file nativename [file join $::workingDir etc xlsx xlsx_data xl worksheets sheet1.xml]]  w+]
    puts $fp [$::mkxlsx::worksheetDoc asXML -xmlDeclaration 1]
    close $fp 

    destroy ::mkxlsx::sharedStringsDoc
    destroy ::mkxlsx::worksheetDoc
    destroy ::mkxlsx::styleSheetDoc
    destroy ::mkxlsx::sharedStringsRoot
    destroy ::mkxlsx::worksheetRoot
    destroy ::mkxlsx::styleSheetRoot
    destroy ::mkxlsx::defaultHeight
    destroy ::mkxlsx::defaultWidth
    destroy ::mkxlsx::boldStyle
    destroy ::mkxlsx::headerStyle
    destroy ::mkxlsx::headerFilledStyle
    destroy ::mkxlsx::textStyle
    destroy ::mkxlsx::dateStyle
    destroy ::mkxlsx::emptyStyle
    destroy ::mkxlsx::maxCol
    destroy ::mkxlsx::worksheet
    destroy ::mkxlsx::rows
    destroy ::mkxlsx::cols
    destroy ::mkxlsx::styles
    destroy ::mkxlsx::fonts
    destroy ::mkxlsx::borders
    destroy ::mkxlsx::fills
}

# ### ### ### ######### ######### #########
## Ready
package provide ::mkxlsx 1.0.0
