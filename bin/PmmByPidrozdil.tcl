encoding system utf-8

namespace eval ::pmm {

    proc getByPidrozdilAction {} \
    {
        set pmmDbRecordSet [::pmm::mysql::runQuery "select a.*, b.entity_name, b.entity_type from pmm_2024 a, dict_pmm b where a.pmm = b.entity_id and 
                            pidrozdil = [lindex [split $::currentPidrozdil . ] 0]
                            order by pmm, date_operation, doc_name"]
        
        ::mkxlsx::xlsxInit

        set myBorder [::mkxlsx::createBorder left thin right thin top thin bottom thin]
        set headerNumStyle [::mkxlsx::createStyle fontId [::mkxlsx::createFont sz 10] horizontal center borderId $myBorder]
        set tableNumStyle [::mkxlsx::createStyle horizontal right borderId $myBorder]
        set fillRedStyle [::mkxlsx::createStyle fillId [::mkxlsx::createFill patternType "solid" bgColor "FFFF0000" fgColor "00FF0000"] vartical "top" horizontal "center"]

        ::mkxlsx::merge "A1:K1"
        ::mkxlsx::addCell A 1 [lindex [split $::currentPidrozdil . ] 1] $::mkxlsx::boldStyle

        ::mkxlsx::merge "A2:A3"
        ::mkxlsx::merge "B2:B3"
        ::mkxlsx::merge "C2:C3"
        ::mkxlsx::merge "D2:D3"
        ::mkxlsx::merge "E2:E3"
        ::mkxlsx::merge "F2:F3"
        ::mkxlsx::merge "G2:G3"
        ::mkxlsx::merge "H2:I2"
 
        ::mkxlsx::addCell A 2 "№" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell B 2 "Дата операції" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell C 2 "Документ" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell D 2 "№ документа" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell E 2 "Дата документа" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell F 2 "Постач./Одержувач" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell G 2 "Назва ПММ" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell H 2 "Прихід" $::mkxlsx::headerFilledStyle

        ::mkxlsx::addCell H 3 "л." $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell I 3 "кг." $::mkxlsx::headerFilledStyle

        ::mkxlsx::addCell A 4 "1" $headerNumStyle
        ::mkxlsx::addCell B 4 "2" $headerNumStyle
        ::mkxlsx::addCell C 4 "3" $headerNumStyle
        ::mkxlsx::addCell D 4 "4" $headerNumStyle
        ::mkxlsx::addCell E 4 "5" $headerNumStyle
        ::mkxlsx::addCell F 4 "6" $headerNumStyle
        ::mkxlsx::addCell G 4 "7" $headerNumStyle
        ::mkxlsx::addCell H 4 "8" $headerNumStyle
        ::mkxlsx::addCell I 4 "9" $headerNumStyle

        if {$::show_minus != 1} {
            ::mkxlsx::merge "J2:K2"
            ::mkxlsx::addCell J 2 "Розхід" $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell J 3 "л." $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell K 3 "кг." $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell J 4 "10" $headerNumStyle
            ::mkxlsx::addCell K 4 "11" $headerNumStyle
        }
        
        set rn 5
        set index 1
        set p_l [set p_k [set r_l [set r_k 0]]]
        set pmm_id [lindex [lindex $pmmDbRecordSet 0] 13]

        foreach row $pmmDbRecordSet {
            foreach {entity_id date_operation doc_name doc_nomer doc_date postach in_l in_k out_l out_k zal_l zal_k pidrozdil pmm date_crea date_modify pmm_name pmm_type} $row {
                if {$pmm_id != $pmm} {
                    ::mkxlsx::merge "A$rn:G$rn"
                    ::mkxlsx::addCell A $rn "Залишок" $::mkxlsx::headerFilledStyle
                    if {$::show_minus == 1 } {
                        ::mkxlsx::addCell H $rn [expr {$p_l - $r_l} ] [expr {($p_l - $r_l) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
                        ::mkxlsx::addCell I $rn [expr {$p_k - $r_k}] [expr {($p_k - $r_k) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
                    } else {
                        ::mkxlsx::addCell H $rn "л." $::mkxlsx::headerFilledStyle
                        ::mkxlsx::addCell I $rn [expr {$p_l - $r_l} ] [expr {($p_l - $r_l) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
                        ::mkxlsx::addCell J $rn "кг." $::mkxlsx::headerFilledStyle
                        ::mkxlsx::addCell K $rn [expr {$p_k - $r_k}] [expr {($p_k - $r_k) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
                    }
                    set p_l [set p_k [set r_l [set r_k 0]]]
                    set pmm_id $pmm
                    incr rn
                }
                ::mkxlsx::addCell A $rn $index $::mkxlsx::dateStyle
                ::mkxlsx::addCell B $rn [clock format [clock scan $date_operation -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                ::mkxlsx::addCell C $rn $doc_name $::mkxlsx::textStyle
                ::mkxlsx::addCell D $rn $doc_nomer $::mkxlsx::dateStyle
                if {$doc_date != "0000-00-00"} {
                    ::mkxlsx::addCell E $rn [clock format [clock scan $doc_date -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                } else {
                    ::mkxlsx::addCell E $rn "" $::mkxlsx::dateStyle
                }
                
                ::mkxlsx::addCell F $rn $postach $::mkxlsx::textStyle
                ::mkxlsx::addCell G $rn $pmm_name $::mkxlsx::textStyle
                
                if {$::show_minus == 1 } {
                    ::mkxlsx::addCell H $rn [expr {$in_l - $out_l}] $tableNumStyle n
                    ::mkxlsx::addCell I $rn [expr {$in_k - $out_k}] $tableNumStyle n
                } else {
                    ::mkxlsx::addCell H $rn $in_l $tableNumStyle n
                    ::mkxlsx::addCell I $rn $in_k $tableNumStyle n
                    ::mkxlsx::addCell J $rn $out_l $tableNumStyle n
                    ::mkxlsx::addCell K $rn $out_k $tableNumStyle n
                }
            }
            incr rn 
            incr index
            set p_l [expr {$p_l + $in_l}]
            set p_k [expr {$p_k + $in_k}]
            set r_l [expr {$r_l + $out_l}]
            set r_k [expr {$r_k + $out_k}]
        }

        ::mkxlsx::merge "A$rn:G$rn"
        ::mkxlsx::addCell A $rn "Залишок" $::mkxlsx::headerFilledStyle
        if {$::show_minus == 1 } {
            ::mkxlsx::addCell H $rn [expr {$p_l - $r_l} ] [expr {($p_l - $r_l) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
            ::mkxlsx::addCell I $rn [expr {$p_k - $r_k}] [expr {($p_k - $r_k) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
        } else {
            ::mkxlsx::addCell H $rn "л." $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell I $rn [expr {$p_l - $r_l} ] [expr {($p_l - $r_l) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
            ::mkxlsx::addCell J $rn "кг." $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell K $rn [expr {$p_k - $r_k}] [expr {($p_k - $r_k) < 0} ? $fillRedStyle : $::mkxlsx::headerFilledStyle] n
        }

        ::mkxlsx::setRowHeight 4 12
        
        ::mkxlsx::setColWidth 1 5
        ::mkxlsx::setColWidth 2 10
        ::mkxlsx::setColWidth 3 15
        ::mkxlsx::setColWidth 4 15
        ::mkxlsx::setColWidth 5 10
        ::mkxlsx::setColWidth 6 15
        ::mkxlsx::setColWidth 7 25
        ::mkxlsx::setColWidth 8 10
        ::mkxlsx::setColWidth 9 10
        ::mkxlsx::setColWidth 10 10
        ::mkxlsx::setColWidth 11 10
        ::mkxlsx::writeXml

        set file_name [::uuid::uuid generate].xlsx

        ::mkxlsx::mkzip [file nativename [file join $::workingDir out $file_name]] -directory [file nativename [file join $::workingDir etc xlsx xlsx_data]]
        catch { [exec $::explorer [file nativename [file join $::workingDir out $file_name]] & ] errorMessage }
    }
}