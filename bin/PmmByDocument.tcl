encoding system utf-8

namespace eval ::pmm {

    proc getByDocumentAction {} \
    {
        set documentsRecordSet [::pmm::mysql::runQuery "select distinct postach, doc_name, doc_nomer, doc_date from pmm_2024 where doc_name != '' order by doc_name, date_operation"]
        lappend documentList [list "Виберіть документ..."]
        set i 1
        foreach record $documentsRecordSet {
            foreach {post name nom date} $record {
                if {$date == "0000-00-00"} {
                    set date ""
                } else {
                    set date [clock format [clock scan $date -format "%Y-%m-%d"] -format "%d.%m.%Y"]
                }
                lappend documentList [list "$name - $date - $nom - $post"]
            }
            incr i
            if {[expr {$i % 20}] == 0} {
                lappend documentList [list "|"]
            }
        }
        set ::doctblcols {0 {entity_id} left  0 {Дата} center 0 {Документ} left 0 {№} left 0 {від} center 0 {Постач/одерж} left \
            0 {Підрозділ} left 0 {ПММ} left 0 {Прихід л.} right 0 {Прихід кг.} right 0 {Розхід л. } right 0 {Розхід кг.} right }

        set ::docForm .docForm
        apave::APave create ::docPave

        ::docPave makeWindow .docForm.fra "Вибір по документу"
        wm geometry .docForm "800x400+0+0"

        set content {
            {toolMenu - - - - {pack -side top} {-array { 
                opcDocument {::searchDocument $documentList { -width 50} {} -command ::pmm::fillDocumentTable -tooltip "Виберіть Документ"} 
            }}}
            {fraDocTable - - - - {pack -side top -fill both -expand 1 -padx 5 -pady 3}}
            {fraDocTable.TblDocMain - - - - {pack -side left -fill both -expand 1} {-h 7 -lbxsel but -columns {$::doctblcols}}}
            {fraDocTable.sbv fraDocTable.tblDocMain L - - {pack -side left -after %w}}
            {fraDocTable.sbh fraDocTable.tblDocMain T - - {pack -side left -before %w}}
        }
        ::docPave paveWindow .docForm.fra $content 
        set ::dbDocTable [::docPave TblDocMain]

        $::dbDocTable columnconfigure 0 -name por_nom -maxwidth 0 -hide yes
        $::dbDocTable columnconfigure 1 -name rec_date -maxwidth 8
        $::dbDocTable columnconfigure 2 -name doc_name -maxwidth 0
        $::dbDocTable columnconfigure 3 -name doc_nom -maxwidth 6
        $::dbDocTable columnconfigure 4 -name doc_date -maxwidth 8 
        $::dbDocTable columnconfigure 5 -name otrim -maxwidth 0
        $::dbDocTable columnconfigure 6 -name pidr -maxwidth 8 
        $::dbDocTable columnconfigure 7 -name pmm -maxwidth 0
        $::dbDocTable columnconfigure 8 -name pryh_l -maxwidth 8
        $::dbDocTable columnconfigure 9 -name pryh_k -maxwidth 8

        wm transient .docForm .mainFrame
        set ::done [::docPave showModal $::docForm]
        destroy $::docForm
        ::docPave destroy
        
        return
    }

    proc fillDocumentTable {} \
    {
        $::dbDocTable columnconfigure 10 -name rozh_l -maxwidth 8 -hide [expr {$::show_minus == 1 ? yes : no}]
        $::dbDocTable columnconfigure 11 -name rozh_k -maxwidth 8 -hide [expr {$::show_minus == 1 ? yes : no}]

        if {[lsearch [split $::searchDocument " "] "Виберіть"] != -1 || [lsearch [split $::searchDocument " "] "Початкові"] != -1 || [lsearch [split $::searchDocument " "] "ЗВІРЕНО"] != -1} {
            return
        }

        foreach {name date nomer postach} [split $::searchDocument "-"] {
            set postach [string trim $postach]
            set doc_name [string trim $name]
            set doc_nomer [string trim $nomer]
            set doc_date [clock format [clock scan [string trim $date] -format "%d.%m.%Y"] -format "%Y-%m-%d"]
        }

        set addWhere " postach = \"$postach\" and doc_name = \"$doc_name\" and doc_nomer = \"$doc_nomer\" and doc_date = \"$doc_date\" "

        set docDBRecordSet [::pmm::mysql::runQuery "select a.*, b.entity_name, c.entity_name from pmm_2024 a, dict_pidrozdil b, dict_pmm c where a.pidrozdil=b.entity_id and a.pmm=c.entity_id and $addWhere order by pmm, pidrozdil, date_operation"]
        $::dbDocTable delete 0 end

        set pmm_name [lindex [lindex $docDBRecordSet 0] 15]
        set sum_in_l [set sum_in_k [set sum_out_l [set sum_out_k 0]]]
        foreach currow $docDBRecordSet {
          foreach {entity_id date_operation doc_name doc_nomer doc_date postach in_l in_k out_l out_k zal_l zal_k pidrozdil pmm date_crea date_modify str_pidr str_pmm} $currow {

            if {$pmm_name != $str_pmm} {
                $::dbDocTable insert end [list "" "" "" "" "" "" "Залишок по" $pmm_name "[expr {$sum_in_l - $sum_out_l}] л." "[expr {$sum_in_k - $sum_out_k}] кг." "" ""]
                $::dbDocTable rowconfigure end -bg "#dee2e6" 
                set sum_in_l [set sum_in_k [set sum_out_l [set sum_out_k 0]]]
                set pmm_name $str_pmm
            }

            if {[string trim $date_operation] !="0000-00-00" } {
                set date_operation [clock format [clock scan $date_operation -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set date_operation ""
            }
            if {[string trim $doc_date] !="0000-00-00" } {
                set doc_date [clock format [clock scan $doc_date -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set doc_date ""
            }

            if {$::show_minus ==1} {
                set pryh_l [expr {$in_l - $out_l}]
                set pryh_k [expr {$in_k - $out_k}]
            } else {
                set pryh_l $in_l
                set pryh_k $in_k
            }

            $::dbDocTable insert end [list $entity_id $date_operation $doc_name $doc_nomer $doc_date $postach $str_pidr $str_pmm $pryh_l $pryh_k $out_l $out_k]
            
            set sum_in_l [expr {$sum_in_l + $in_l}]
            set sum_in_k [expr {$sum_in_k + $in_k}]
            set sum_out_l [expr {$sum_out_l + $out_l}]
            set sum_out_k [expr {$sum_out_k + $out_k}]

            if {$doc_name == "Початкові значення"} {
                $::dbDocTable rowconfigure end -bg "#ba0000"
            }
            
            if {$doc_name == "Зведена відомість"} {
                $::dbDocTable rowconfigure end -bg "#00b0f0"
            }

            if {[string toupper $doc_name] == "ЗВІРЕНО"} {
                $::dbDocTable rowconfigure end -bg "#00b050"
            }

            if {[string toupper $postach] == "ПІДРОЗДІЛИ"} {
                $::dbDocTable rowconfigure end -bg "#7030a0"
            }
          }
        }
        $::dbDocTable insert end [list "" "" "" "" "" "" "Залишок по" $pmm_name "[expr {$sum_in_l - $sum_out_l}] л." "[expr {$sum_in_k - $sum_out_k}] кг." "" ""]
        $::dbDocTable rowconfigure end -bg "#dee2e6" 

        $::dbDocTable yview moveto 1.0
        return
    }
}
