encoding system utf-8

namespace eval ::pmm {

    proc getByPmmAction {} \
    {        
        set pmmDbRecordSet [::pmm::mysql::runQuery "select a.*, b.entity_name from pmm_2024 a, dict_pidrozdil b where a.pidrozdil = b.entity_id and 
                            pmm = [lindex [split $::currentPMM . ] 0]
                            order by date_operation, doc_name, doc_nomer, doc_date, postach, pidrozdil"]
        
        set idList [::pmm::mysql::runQuery "select distinct pidrozdil from pmm_2024 where (in_k > 0 or out_k > 0)
                            and pmm = [lindex [split $::currentPMM . ] 0] and pidrozdil != 26"]
        
        set pidrozdilList [::pmm::mysql::runQuery "select * from dict_pidrozdil where entity_id in ([join $idList ,])"]
        set pmmName [::pmm::mysql::runQuery "select * from dict_pmm where entity_id = [lindex [split $::currentPMM . ] 0]"]
        
        ::mkxlsx::xlsxInit

        ::mkxlsx::setTabName [lindex [lindex $pmmName 0] 1]

        set myBorder [::mkxlsx::createBorder left thin right thin top thin bottom thin]
        set headerNumStyle [::mkxlsx::createStyle fontId [::mkxlsx::createFont sz 10] horizontal center borderId $myBorder]
        set headerStyle [::mkxlsx::createStyle fontId [::mkxlsx::createFont sz 10 b true] fillId [::mkxlsx::createFill patternType "solid" bgColor "FF000000" fgColor "FFFFFFFF"] \
             vertical "top" horizontal "center" borderId $myBorder wrapText true]
        set tableYellowTextStyle [::mkxlsx::createStyle numFmtId 165 fontId [::mkxlsx::createFont sz 10] fillId [::mkxlsx::createFill patternType "solid" bgColor "FF000000" fgColor "FFFFE699"] \
             vertical "top" horizontal "center" borderId $myBorder wrapText true xfId 0]
        set tableSumTextStyle [::mkxlsx::createStyle numFmtId 165 fontId [::mkxlsx::createFont sz 8] fillId [::mkxlsx::createFill patternType "solid" bgColor "FF000000" fgColor "FFFFFFFF"] \
             vertical "top" horizontal "center" borderId $myBorder wrapText true xfId 0]
        set tableBlueTextStyle [::mkxlsx::createStyle numFmtId 165 fontId [::mkxlsx::createFont sz 10] fillId [::mkxlsx::createFill patternType "solid" bgColor "FF000000" fgColor "FF00B0F0"] \
             vertical "top" horizontal "center" borderId $myBorder wrapText true xfId 0]
        set tableFioletTextStyle [::mkxlsx::createStyle numFmtId 165 fontId [::mkxlsx::createFont sz 10] fillId [::mkxlsx::createFill patternType "solid" bgColor "FF000000" fgColor "FF7030A0"] \
             vertical "top" horizontal "center" borderId $myBorder wrapText true xfId 0]

        set tableTextStyle $tableYellowTextStyle

        foreach col {A B C D E M} {
            ::mkxlsx::merge "${col}1:${col}3"
        }
        ::mkxlsx::merge "F1:H2"
        ::mkxlsx::merge "I1:L1"
        ::mkxlsx::merge "I2:K2"

        set firstCol [set currCol 13]

        ::mkxlsx::addCell A 1 "Дата запису" $headerStyle
        ::mkxlsx::addCell B 1 "Найменування документа" $headerStyle
        ::mkxlsx::addCell C 1 "№ документа" $headerStyle
        ::mkxlsx::addCell D 1 "Дата документа" $headerStyle
        ::mkxlsx::addCell E 1 "Постачальник (одержувач)" $headerStyle
        ::mkxlsx::addCell F 1 "За в/ч А7015" $headerStyle
        ::mkxlsx::addCell I 1 "Склад" $headerStyle
        ::mkxlsx::addCell M 1 "Всього за підрозділи" $headerStyle

        ::mkxlsx::addCell I 2 "Склад В/Ч" $headerStyle
        ::mkxlsx::addCell L 2 "ПСП" $headerStyle

        ::mkxlsx::addCell F 3 "Надійшло" $headerStyle
        ::mkxlsx::addCell G 3 "Вибуло" $headerStyle
        ::mkxlsx::addCell H 3 "Усього" $headerStyle

        ::mkxlsx::addCell I 3 "Надійшло" $headerStyle
        ::mkxlsx::addCell J 3 "Вибуло" $headerStyle
        ::mkxlsx::addCell K 3 "Усього" $headerStyle
        ::mkxlsx::addCell L 3 "Усього" $headerStyle

        foreach row $pidrozdilList {
            incr currCol
            foreach {id name} $row {
                ::mkxlsx::setColWidth $currCol 7
                ::mkxlsx::merge "[::pmm::sysutils::getCol $currCol]1:[::pmm::sysutils::getCol [expr {$currCol + 1}]]2"
                ::mkxlsx::addCell [::pmm::sysutils::getCol $currCol] 1 $name $headerStyle
                ::mkxlsx::addCell [::pmm::sysutils::getCol $currCol] 3 "л." $headerStyle
                lappend cells [::pmm::sysutils::getCol $currCol]
                dict set departments [::pmm::sysutils::getCol $currCol] name $name
                dict set departments [::pmm::sysutils::getCol $currCol] col [::pmm::sysutils::getCol $currCol]
                dict set departments [::pmm::sysutils::getCol $currCol] value "1"
                incr currCol
                ::mkxlsx::setColWidth $currCol 7
                ::mkxlsx::addCell [::pmm::sysutils::getCol $currCol] 3 "кг." $headerStyle
                lappend cells [::pmm::sysutils::getCol $currCol]
                dict set departments [::pmm::sysutils::getCol $currCol] name $name
                dict set departments [::pmm::sysutils::getCol $currCol] col [::pmm::sysutils::getCol $currCol] 
                dict set departments [::pmm::sysutils::getCol $currCol] value "2"
                lappend suma [::pmm::sysutils::getCol $currCol]
            }
        }

        ::mkxlsx::addRow 4 $tableSumTextStyle

        set rn 5
        set index 1
        set document "[lindex [lindex $pmmDbRecordSet 0] 1][lindex [lindex $pmmDbRecordSet 0] 2][lindex [lindex $pmmDbRecordSet 0] 3][lindex [lindex $pmmDbRecordSet 0] 4][lindex [lindex $pmmDbRecordSet 0] 5]"

        foreach row $pmmDbRecordSet {
            foreach {entity_id date_operation doc_name doc_nomer doc_date postach in_l in_k out_l out_k zal_l zal_k pidrozdil pmm date_crea date_modify p_name} $row {
                set zal_l [expr {int($in_l - $out_l)}]
                set zal_k [expr {int($in_k - $out_k)}]
                set in_l [expr {int($in_l)}]
                set in_k [expr {int($in_k)}]
                set out_l [expr {int($out_l)}] 
                set out_k [expr {int($out_k)}] 
        
                ::mkxlsx::setRowHeight $rn 15

                if {$document != "$date_operation$doc_name$doc_nomer$doc_date$postach"} {
                    set document "$date_operation$doc_name$doc_nomer$doc_date$postach"
                    incr rn 2
                    
                    if {$postach == "Донесення"} {
                        set tableTextStyle $tableBlueTextStyle
                    } elseif {$postach == "Підрозділи"} {
                        set tableTextStyle $tableFioletTextStyle
                    } else {
                        set tableTextStyle $tableYellowTextStyle
                    }

                    ::mkxlsx::addRow $rn $tableTextStyle
                    ::mkxlsx::addRow [expr {$rn + 1}] $tableSumTextStyle
                    ::mkxlsx::setRowHeight [expr {$rn + 1}] 15

                    ::mkxlsx::addCell F [expr {$rn + 1}] "" $tableSumTextStyle n "F[expr {$rn - 1}]+F$rn"
                    ::mkxlsx::addCell G [expr {$rn + 1}] "" $tableSumTextStyle n "G[expr {$rn - 1}]+G$rn"
                    ::mkxlsx::addCell H [expr {$rn + 1}] "" $tableSumTextStyle n "H[expr {$rn - 1}]+H$rn"
                    ::mkxlsx::addCell K [expr {$rn + 1}] "" $tableSumTextStyle n "K[expr {$rn - 1}]+I$rn-J$rn"
                    ::mkxlsx::addCell L [expr {$rn + 1}] "" $tableSumTextStyle n "L[expr {$rn - 1}]+L$rn"
                    ::mkxlsx::addCell M [expr {$rn + 1}] "" $tableSumTextStyle n "M[expr {$rn - 1}]+M$rn"

                    foreach cell $cells {
                        ::mkxlsx::addCell $cell [expr {$rn + 1}] "" $tableSumTextStyle n "$cell[expr {$rn - 1}]+$cell$rn"
                    }
                }
                ::mkxlsx::addCell A $rn [clock format [clock scan $date_operation -format "%Y-%m-%d"] -format "%d.%m.%Y"] $tableTextStyle
                ::mkxlsx::addCell B $rn $doc_name $tableTextStyle
                ::mkxlsx::addCell C $rn $doc_nomer $tableTextStyle
                if {$doc_date != "0000-00-00"} {
                    ::mkxlsx::addCell D $rn [clock format [clock scan $doc_date -format "%Y-%m-%d"] -format "%d.%m.%Y"] $tableTextStyle
                } else {
                    ::mkxlsx::addCell D $rn "" $tableTextStyle
                }
                ::mkxlsx::addCell E $rn $postach $tableTextStyle
                
                ::mkxlsx::addCell F $rn "" $tableTextStyle n
                ::mkxlsx::addCell G $rn "" $tableTextStyle n
                ::mkxlsx::addCell H $rn "" $tableTextStyle n "K$rn+L$rn+M$rn"
                ::mkxlsx::addCell F [expr {$rn + 1}] "" $tableSumTextStyle n "F[expr {$rn - 1}]+F$rn"
                ::mkxlsx::addCell G [expr {$rn + 1}] "" $tableSumTextStyle n "G[expr {$rn - 1}]+G$rn"
                ::mkxlsx::addCell H [expr {$rn + 1}] "" $tableSumTextStyle n "H[expr {$rn - 1}]+H$rn"

                if {$pidrozdil == 26} {
                    ::mkxlsx::addCell I $rn $in_k $tableTextStyle n
                    ::mkxlsx::addCell J $rn $out_k $tableTextStyle n
                    ::mkxlsx::addCell K $rn "" $tableTextStyle n "K[expr {$rn - 1}]+I$rn-J$rn"
                    ::mkxlsx::addCell K [expr {$rn + 1}] "" $tableSumTextStyle n "K$rn+I[expr {$rn + 1}]-J[expr {$rn + 1}]"
                } else {
                    ::mkxlsx::addCell I $rn "" $tableTextStyle n
                    ::mkxlsx::addCell J $rn "" $tableTextStyle n
                    ::mkxlsx::addCell K $rn "" $tableTextStyle n "K[expr {$rn - 1}]+I$rn-J$rn"
                    ::mkxlsx::addCell K [expr {$rn + 1}] "" $tableSumTextStyle n "K$rn+I[expr {$rn + 1}]-J[expr {$rn + 1}]"
                }

                ::mkxlsx::addCell L $rn "" $tableTextStyle n
                ::mkxlsx::addCell L [expr {$rn + 1}] "" $tableSumTextStyle n "L[expr {$rn - 1}]+L$rn"

                ::mkxlsx::addCell M $rn "" $tableTextStyle n "SUM([join $suma $rn,]$rn)"
                ::mkxlsx::addCell M [expr {$rn + 1}] "" $tableSumTextStyle n "M[expr {$rn - 1}]+M$rn"
                
                foreach cell $cells {
                    if {[dict get $departments $cell name] == $p_name } {
                        ::mkxlsx::addCell [dict get $departments $cell col] $rn [expr {[dict get $departments $cell value] == "1" ? $zal_l : $zal_k}] $tableTextStyle n
                        ::mkxlsx::addCell [dict get $departments $cell col] [expr {$rn + 1}] "" $tableSumTextStyle n "[dict get $departments $cell col][expr {$rn - 1}]+[dict get $departments $cell col]$rn"
                    }
                }
            }
        }
        ::mkxlsx::setColWidth 1 7
        ::mkxlsx::setColWidth 2 10
        ::mkxlsx::setColWidth 3 10
        ::mkxlsx::setColWidth 4 7
        ::mkxlsx::setColWidth 5 10
        ::mkxlsx::setColWidth 6 7
        ::mkxlsx::setColWidth 7 7
        ::mkxlsx::setColWidth 8 7
        ::mkxlsx::setColWidth 9 7
        ::mkxlsx::setColWidth 10 7
        ::mkxlsx::setColWidth 11 7
        ::mkxlsx::setColWidth 12 5
        ::mkxlsx::setColWidth 13 7
        ::mkxlsx::splitSheet 13 3 [::pmm::sysutils::getCol 14]
        ::mkxlsx::writeXml
        set file_name [::uuid::uuid generate].xlsx

        ::mkxlsx::mkzip [file nativename [file join $::workingDir out $file_name]] -directory [file nativename [file join $::workingDir etc xlsx xlsx_data]]
        catch { [exec $::explorer [file nativename [file join $::workingDir out $file_name]] & ] errorMessage }
    }
}

