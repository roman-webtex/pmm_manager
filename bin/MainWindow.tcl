encoding system utf-8

namespace eval ::pmm {
    
    proc buildMainWindow {} {
        set ::searchVariable ""
        set ::mainTitle "Облік ПММ"
        set ::db_table_name "pmm_2024"
        set ::workingYear [clock format [clock seconds] -format "%Y"]
        set ::empty_date "0000-00-00"
        
        apave::APave create ::pave
        set ::checkedImg   [image create photo tbl_checkedImg   -data [apave::iconData warn]]
        
        foreach {i icon} {0 add 1 change 2 warn 3 OpenFile 4 folder} {
            image create photo Img$i -data [apave::iconData $icon]
        }
        
        set ::pidrozdilRecordSet [::pmm::mysql::getRecordSetWithQuery "*" dict_pidrozdil "1 order by entity_id"]
        set ::pmmRecordSet [::pmm::mysql::getRecordSetWithQuery "*" dict_pmm "1 order by entity_id"]
        
        set i 0
        foreach record $::pidrozdilRecordSet {
            foreach {p_id p_name} $record {
                lappend ::pidrozdilList [list "$p_id. $p_name"]
                lappend ::postachList "$p_name"
            }
        }
        foreach record $::pmmRecordSet {
            foreach {p_id p_name p_type} $record {
                lappend ::pmmList [list "$p_id. $p_name"]
                incr i
                if {[expr {$i % 40}] == 0} {
                    lappend ::pmmList [list "|"]
                }
            }
        }


        set ::tblcols {0 {entity_id} left  0 {Дата запису} center 0 {Найменування документа} left 0 {№ документа} left 0 {Дата документа} center 0 {Постачальник/одержувач} left \
                       0 {Прихід л.} right 0 {Прихід кг.} right 0 {Розхід л. } right 0 {Розхід кг.} right 0 {Залишок л.} right 0 {Залишок кг.} right}
        
        toplevel .mainFrame
        wm title .mainFrame $::mainTitle

        wm geometry .mainFrame "$::window_size"

        set content {
            {Menu - - - - - {-array {File "&Файл"  View "&Довідники" Sys "&Система"}} ::pmm::fillMenuAction}
            {toolMenu - - - - {pack -side top} {-array { 
                Img0 {{::pmm::addAction} -tooltip "Додати запис \nCtrl+N"} 
                Img1 {{::pmm::editAction} -tooltip "Редагувати запис \nCtrl+Enter"} sev 5 h_ 1
                OpcPidrozdil {::currentPidrozdil ::pidrozdilList {-width 16} {} -command ::pmm::fillTableAction -tooltip "Виберіть підрозділ"} 
                OpcPmm {::currentPMM ::pmmList {-width 26} {} -command ::pmm::fillTableAction -tooltip "Виберіть тип ПММ"} sev 5 h_ 1
                ChbAll      { -command {::pmm::fillTableAction} -var ::show_minus -t "Від'ємний розхід" -tooltip "Показувати розхід в одній колонці з приходом"} sev 5 h_ 1
            }}}
            {fraSearch - - - - {pack -side top -fill x -expand 0 -padx 5 -pady 3} {-relief flat}}
            {fraSearch.entSearch - - - - {pack -fill x} {-tvar ::searchVariable}}
            {fraTable - - - - {pack -side top -fill both -expand 1 -padx 5 -pady 3}}
            {fraTable.TblMain - - - - {pack -side left -fill both -expand 1} {-h 7 -lbxsel but -columns {$::tblcols}}}
            {fraTable.sbv fraTable.tblMain L - - {pack -side left -after %w}}
            {fraTable.sbh fraTable.tblMain T - - {pack -side left -before %w}}
            {staZalyshkiZagal - - - - {pack -side bottom -padx 5 } { -array { 
                {"За складом ПММ, кг. : " -font {-slant italic -size 10} -anchor w} 60
                {"За військовою частиною, кг.: " -font {-slant italic -size 10} -anchor w} 60 }}}
            {staZalyshkiPidrozdily - - - - {pack -side bottom -padx 5 } { -array { 
                {"Загалом по підрозділах, л. : " -font {-slant italic -size 8} -anchor w} 60
                {"Загалом по підрозділах, кг.: " -font {-slant italic -size 8} -anchor w} 60 }}}
            {staZalyshki - - - - {pack -side bottom -padx 5 } { -array { 
                {"Залишок у підрозділі, л. : " -font {-slant italic -size 8} -anchor w} 60
                {"Залишок у підрозділі, кг.: " -font {-slant italic -size 8} -anchor w} 60 }}}
        }

        ::pave paveWindow .mainFrame $content

        set ::dbTable [::pave TblMain]
        set ::cbPidrozdil [::pave OpcPidrozdil]
        set ::cbPmm [::pave OpcPmm]

        $::dbTable columnconfigure 0 -name por_nom -maxwidth 0 -hide yes
        $::dbTable columnconfigure 1 -name rec_date -maxwidth 0
        $::dbTable columnconfigure 2 -name doc_name -maxwidth 0
        $::dbTable columnconfigure 3 -name doc_nom -maxwidth 0
        $::dbTable columnconfigure 4 -name doc_date -maxwidth 0 
        $::dbTable columnconfigure 5 -name otrim -maxwidth 0
        $::dbTable columnconfigure 6 -name prih_l -maxwidth 8 
        $::dbTable columnconfigure 7 -name prih_k -maxwidth 8
        $::dbTable columnconfigure 10 -name zalish_l -maxwidth 8 -hide yes
        $::dbTable columnconfigure 11 -name zalish_k -maxwidth 8 -hide yes

        focus .mainFrame.fraSearch.entSearch

        wm protocol .mainFrame WM_DELETE_WINDOW {::pmm::prog_exit}
        
        bind .mainFrame <F2> {::pmm::addAction}
        bind .mainFrame <F4> {::pmm::editAction}
        bind .mainFrame <Return> {::pmm::editAction}
        bind .mainFrame <F10> {::pmm::prog_exit}
        bind .mainFrame <Escape> {::pmm::fillTableAction}

        ::pmm::fillTableAction
    }
    
    proc fillTableAction {{id_list 0} {where ""}} \
    {
        $::dbTable columnconfigure 8 -name rozh_l -maxwidth 8 -hide [expr {$::show_minus == 1 ? yes : no}]
        $::dbTable columnconfigure 9 -name rozh_k -maxwidth 8 -hide [expr {$::show_minus == 1 ? yes : no}]
        set addWhere "pidrozdil = [lindex [split $::currentPidrozdil "."] 0] and pmm = [lindex [split $::currentPMM "."] 0]"

        if {$id_list != 0} {
           set addWhere "entity_id in ($id_list) and $addWhere"
        } elseif {$where != ""} {
           set addWhere "$where and $addWhere"
        }

        set currentDBRecordSet [::pmm::mysql::getRecordSetWithQuery "*" $::db_table_name "$addWhere order by date_operation, entity_id"]
        $::dbTable delete 0 end

        foreach currow $currentDBRecordSet {
          foreach {entity_id date_operation doc_name doc_nomer doc_date postach in_l in_k out_l out_k zal_l zal_k pidrozdil pmm date_crea date_modify} $currow {
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

            $::dbTable insert end [list $entity_id $date_operation $doc_name $doc_nomer $doc_date $postach $pryh_l $pryh_k $out_l $out_k $zal_l $zal_k]
            
            if {$doc_name == "Початкові значення"} {
                $::dbTable rowconfigure end -bg "#ba0000"
            }
            
            if {$doc_name == "Зведена відомість"} {
                $::dbTable rowconfigure end -bg "#00b0f0"
            }

            if {[string toupper $doc_name] == "ЗВІРЕНО"} {
                $::dbTable rowconfigure end -bg "#00b050"
            }

            if {[string toupper $postach] == "ПІДРОЗДІЛИ"} {
                $::dbTable rowconfigure end -bg "#7030a0"
            }
          }
        }
        $::dbTable yview moveto 1.0
        wm title .mainFrame "$::mainTitle"

        set ::zalyshok [::pmm::mysql::runQuery "select sum(in_l)-sum(out_l), sum(in_k)-sum(out_k) from $::db_table_name where pidrozdil = [lindex [split $::currentPidrozdil "."] 0] and pmm = [lindex [split $::currentPMM "."] 0]"]
        set ::zalyshokPidrozdily [::pmm::mysql::runQuery "select sum(in_l)-sum(out_l), sum(in_k)-sum(out_k) from $::db_table_name where pmm = [lindex [split $::currentPMM "."] 0] and pidrozdil != 26"]
        set ::zalyshokSklad [::pmm::mysql::runQuery "select sum(in_k)-sum(out_k) from $::db_table_name where pmm = [lindex [split $::currentPMM "."] 0] and pidrozdil = 26"]

        [::pave LabstaZalyshki1] configure -text [lindex [lindex $::zalyshok 0] 0]
        [::pave LabstaZalyshki2] configure -text [lindex [lindex $::zalyshok 0] 1]
        [::pave LabstaZalyshkiPidrozdily1] configure -text [lindex [lindex $::zalyshokPidrozdily 0] 0]
        [::pave LabstaZalyshkiPidrozdily2] configure -text [lindex [lindex $::zalyshokPidrozdily 0] 1]
        [::pave LabstaZalyshkiZagal1] configure -text [lindex [lindex $::zalyshokSklad 0] 0]
        [::pave LabstaZalyshkiZagal2] configure -text [expr {([lindex [lindex $::zalyshokSklad 0] 0] == "" ? 0 : [lindex [lindex $::zalyshokSklad 0] 0]) + ([lindex [lindex $::zalyshokPidrozdily 0] 1] == "" ? 0 : [lindex [lindex $::zalyshokPidrozdily 0] 1])}]

        focus .mainFrame.fraSearch.entSearch
        return
    }

    proc searchAction {} \
    {
        if {[string trim $::searchVariable] == ""} {
            ::pmm::fillTableAction
            return
        }

        $::dbTable setbusycursor

        set searchList [split $::searchVariable ";"]

        set result {}
        set rowCount [$::dbTable size]

        foreach item [$::dbTable get 0 $rowCount] {
            foreach value $item {
                foreach sval $searchList {
                    if {[regexp -nocase -- "[string trim $sval]" $value] > 0} {
                        lappend result [lindex $item 0]
                        lappend result ","
                    }
                }
            }
        }
        
        if {[llength $result] > 0} {
            set ::inSearch 1
            set ::filePath ""
            wm title .mainFrame "Пошук - $::searchVariable"
            $::dbTable restorecursor
            ::pmm::fillTableAction [lrange $result 0 end-1]
        } else {
            tk_messageBox -message "\"$::searchVariable\"" -type ok -detail "По вашому запиту нічого не знайдено" -icon info
            $::dbTable restorecursor
            ::pmm::fillTableAction
        }
        return
    }

    proc setYear {} \
    {
        set ::db_table_name "pmm_[string trim $::table_year]"
        ::pmm::fillTableAction
        return
    }

    proc editAction {} \
    {
        set row [lindex [$::dbTable get [$::dbTable curselection] [$::dbTable curselection]] 0]
        set ::entity_id      [lindex $row 0]
        set ::date_operation [lindex $row 1]
        set ::doc_name       [lindex $row 2]
        set ::doc_nomer      [lindex $row 3]
        set ::doc_date       [lindex $row 4]
        set ::postach        [lindex $row 5]
        set ::in_l           [lindex $row 6]
        set ::in_k           [lindex $row 7]
        set ::out_l          [lindex $row 8]
        set ::out_k          [lindex $row 9]

        ::pmm::editWindowAction

        if {$::done == 1} {
            set ::date_operation [clock format [clock scan $::date_operation -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            if {$::doc_date != ""} {
                set ::doc_date [clock format [clock scan $::doc_date -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::doc_date $::empty_date
            }

            if {$::in_l == ""} {
                set ::in_l 0
            }
            if {$::in_k == ""} {
                set ::in_k 0
            }
            if {$::out_l == ""} {
                set ::out_l 0
            }
            if {$::out_k == ""} {
                set ::out_k 0
            }

            set ::in_l  [string map {"," "."} $::in_l]
            set ::in_k  [string map {"," "."} $::in_k]
            set ::out_l [string map {"," "."} $::out_l]
            set ::out_k [string map {"," "."} $::out_k]
            
            set query "update $::db_table_name set date_operation = \"$::date_operation\", 
                                          doc_name = \"$::doc_name\", 
                                          doc_nomer = \"$::doc_nomer\", 
                                          doc_date = \"$::doc_date\", 
                                          postach = \"$::postach\", 
                                          in_l = $::in_l ,
                                          in_k = $::in_k ,
                                          out_l = $::out_l ,
                                          out_k = $::out_k ,
                                          date_modify = now()
                                          where entity_id = $::entity_id"
            ::pmm::mysql::runQuery $query
        }
        ::pmm::searchAction
        return
    }

    proc addAction {} \
    {
        set ::doc_name [set ::doc_nomer [set ::postach ""]]
        set ::date_operation [set ::doc_date [clock format [clock seconds] -format {%d.%m.%Y}]]
        set ::in_l [set ::in_k [set ::out_l [set ::out_k [set ::zal_l [set ::zal_k 0]]]]]
        
        ::pmm::editWindowAction

        if {$::done == 1} {

            set ::date_operation [clock format [clock scan $::date_operation -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            if {$::doc_date != ""} {
                set ::doc_date [clock format [clock scan $::doc_date -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::doc_date $::empty_date
            }

            if {$::in_l == ""} {
                set ::in_l 0
            }
            if {$::in_k == ""} {
                set ::in_k 0
            }
            if {$::out_l == ""} {
                set ::out_l 0
            }
            if {$::out_k == ""} {
                set ::out_k 0
            }

            set ::in_l  [string map {"," "."} $::in_l]
            set ::in_k  [string map {"," "."} $::in_k]
            set ::out_l [string map {"," "."} $::out_l]
            set ::out_k [string map {"," "."} $::out_k]
            set ::zal_l 0
            set ::zal_k 0
            
            set in_pidrozdil [lindex [split $::currentPidrozdil "."] 0]
            set in_pmm [lindex [split $::currentPMM "."] 0]

            set query "insert into $::db_table_name values (null, \"$::date_operation\", \"$::doc_name\", \"$::doc_nomer\", \"$::doc_date\", \"$::postach\", 
                      $::in_l, $::in_k, $::out_l, $::out_k, $::zal_l, $::zal_k, $in_pidrozdil, $in_pmm, now(), now())"
            ::pmm::mysql::runQuery $query

            set p_in [::pmm::mysql::runQuery "select entity_id from dict_pidrozdil where entity_name = \"$::postach\" "]

            if {$::in_k == 0 && $::out_k != 0 && $p_in != ""} {
                set result [tk_dialog .dial "Додати прихід" "Додати прихід ($::out_l л. $::out_k кг.) для $::postach?" questhead 1 Так Ні]
                if { $result == 0} {
                    set query "insert into $::db_table_name values (null, \"$::date_operation\", \"$::doc_name\", \"$::doc_nomer\", \"$::doc_date\", \"$::postach\", 
                              $::out_l, $::out_k, 0, 0, 0, 0, $p_in, $in_pmm)"
                    ::pmm::mysql::runQuery $query
                }
            }
        }
        ::pmm::fillTableAction
        return
    }

    proc editWindowAction {} \
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

        set ::editForm .addForm
        apave::APave create ::editPave

        ::editPave makeWindow .addForm.fra [expr {($::doc_name eq "") ? "Новий документ $::currentPMM" : "Редагування: \[ $::doc_name \] $::currentPMM"}]

            
        set content {
            {toolMenu - - - - {pack -side top} {-array { 
                opcDocument {::currentDocument $documentList { -width 50} {} -command ::pmm::applyDocument -tooltip "Виберіть Документ"} 
            }}}
            {fraDocument - - - - {pack -side top -padx 5 -pady 5 -fill both -expand 1}}
            {fraDocument.labDateOperation - - 1 1 {-st we} {-t "Дата операції :" -font {-size $::fSize} }}
            {fraDocument.datDateOperation fraDocument.labDateOperation L 1 1 {-st we} {-tvar ::date_operation -title {Дата операції} -dateformat %d.%m.%Y -width 10 -font {-size $::fSize} }}
            {fraDocument.labPostach fraDocument.datDateOperation L 1 1 {-st we} {-t "Постачальник (одержувач) :" -font {-size $::fSize} }}
            {fraDocument.fcoPostach fraDocument.labPostach L 1 1 {-st we} {-tvar ::postach -values { -list { $::postachList} } -tooltip "Вибір постачальника/одержувача" -font {-size $::fSize} }}
            {fraDocument.labDocName fraDocument.labDateOperation T 1 1 {-st we} {-t "Назва документа :" -font {-size $::fSize} }}
            {fraDocument.entDocName fraDocument.labDocName L 1 1 {-st we} {-tvar ::doc_name -tooltip "$::doc_name" -font {-size $::fSize} }}
            {fraDocument.labNomDoc fraDocument.entDocName L 1 1 {-st we} {-t "№ документа :" -font {-size $::fSize} }}
            {fraDocument.entNomDoc fraDocument.labNomDoc L 1 1 {-st we} {-tvar ::doc_nomer -tooltip "$::doc_nomer"}}
            {fraDocument.labDocDate fraDocument.entNomDoc L 1 1 {-st we} {-t " від :" -font {-size $::fSize} }}
            {fraDocument.datDocDate fraDocument.labDocDate L 1 1 {-st we} {-tvar ::doc_date -title {Дата документа} -dateformat %d.%m.%Y -width 10 -font {-size $::fSize} }}
            {seh2 fraDocument T 1 1 {pack -fill x}}
            {fraValues - - 1 6 {pack -side top -padx 5 -pady 5 -fill both -expand 1}}
            {fraValues.labPryhid - - 1 2 {-st we} {-t "Прихід :" -font {-size $::fSize} }}
            {fraValues.entPryhidL fraValues.labPryhid L 1 1 {-st we} {-tvar ::in_l -tooltip "Прихід, літри" -font {-size $::fSize} -justify right}}
            {fraValues.labPryhidL fraValues.entPryhidL L 1 1 {-st we} {-t "л." -font {-size $::fSize} }}
            {fraValues.entPryhidK fraValues.labPryhidL L 1 1 {-st we} {-tvar ::in_k -tooltip "Прихід, кілограми" -font {-size $::fSize} -justify right}}
            {fraValues.labPryhidK fraValues.entPryhidK L 1 1 {-st we} {-t "кг." -font {-size $::fSize} }}
            {fraValues.labRozhid fraValues.labPryhid T 1 2 {-st we} {-t "Розхід :" -font {-size $::fSize} }}
            {fraValues.entRozhidL fraValues.labRozhid L 1 1 {-st we} {-tvar ::out_l -tooltip "Розхід, літри" -font {-size $::fSize} -justify right }}
            {fraValues.labRozhidL fraValues.entRozhidL L 1 1 {-st we} {-t "л." -font {-size $::fSize} }}
            {fraValues.entRozhidK fraValues.labRozhidL L 1 1 {-st we} {-tvar ::out_k -tooltip "Розхід, кілограми" -font {-size $::fSize} -justify right }}
            {fraValues.labRozhidK fraValues.entRozhidK L 1 1 {-st we} {-t "кг." -font {-size $::fSize} }}
            {seh1 fraValues T 1 1 {pack -fill x}}
            {fraButton - - - - {pack -side top -fill x -expand 0}}
            {.butSave .butSaveNext L 1 1 {pack -side right -padx 5 -pady 3} {-t "Зберегти" -com ::pmm::saveCheck}}
            {.butSaveNext .butCancel L 1 1 {pack -side right -padx 5 -pady 3} {-t "Зберегти та додати" -com ::pmm::saveAndAdd}}
            {.butCancel - - 1 1 {pack -side right -padx 5 -pady 3} {-t "Відміна" -com "::editPave res $::editForm 0"}}
        }
        ::editPave paveWindow .addForm.fra $content 

        wm transient .addForm .mainFrame
        set ::done [::editPave showModal $::editForm -focus .addForm.fra.fraInner.entName]
        destroy $::editForm
        ::editPave destroy
        return
    }

    proc saveCheck {} \
    {
        if {$::date_operation != "" && ($::doc_name == "" || $::doc_nomer =="" || $::doc_date == "") && $::doc_name != "Початкові значення"} {
            tk_messageBox -message "Запис містить помилкі!" -detail "Не заповнено документ, або його номер та дату"  -icon error
        } else {
            ::editPave res $::editForm 1
        }
    }
    
    proc getZalyshkiPidrozdil {type} \
    {
        set zalysh [lindex [::pmm::mysql::runQuery "select sum(in_l), sum(in_k), sum(out_l), sum(out_k) from pmm_2024 where 
                                    pidrozdil = [lindex [split $::currentPidrozdil .] 0] and 
                                    pmm in (select entity_id from dict_pmm where entity_type like \"$type\")"] 0]
        
        set fmt "%-*s: %*.1f л., %*.1f кг."

        tk_messageBox -message "$type, [lindex [split $::currentPidrozdil .] 1]" \
                -detail [format $fmt 10 "Залишок" 15 [expr {[lindex $zalysh 0] - [lindex $zalysh 2]}] 15 [expr {[lindex $zalysh 1] - [lindex $zalysh 3]}]]
        return
    }

    proc getZalyshki {type} \
    {
        set zalysh [lindex [::pmm::mysql::runQuery "select sum(in_l), sum(in_k), sum(out_l), sum(out_k) from pmm_2024 where 
                                    pmm in (select entity_id from dict_pmm where entity_type like \"$type\")"] 0]
        
        set fmt "%-*s: %*.1f л., %*.1f кг."

        tk_messageBox -message "$type" \
                -detail [format $fmt 10 "Залишок" 15 [expr {[lindex $zalysh 0] - [lindex $zalysh 2]}] 15 [expr {[lindex $zalysh 1] - [lindex $zalysh 3]}]]
        return
    }
    
    proc getDatesInterval {} \
    {
        apave::APave create ::wndDates
        set ::dateFrame .dateFrame
        set ::dateBegin [clock format [clock seconds] -format "%d.%m.%Y"]
        set ::dateEnd [clock format [clock seconds] -format "%d.%m.%Y"]

        set content {
            {fra  - - - - {-st news -padx 5 -pady 5}}
            {fra.lab1 - - 1 1    {-st es}  {-t "Дата з: "}}
            {fra.dat1 fra.lab1 L 1 9 {-st wes} {-tvar ::dateBegin -title {Початкова дата} -dateformat %d.%m.%Y -width 10 -font {-size $::fSize} }}
            {fra.lab2 fra.lab1 T 1 1 {-st es}  {-t "Дата по: "}}
            {fra.dat2 fra.lab2 L 1 9 {-st wes} {-tvar ::dateEnd -title {Кінцева дата} -dateformat %d.%m.%Y -width 10 -font {-size $::fSize} }}
            {fra.seh1 fra.lab2 T 1 10 }
            {fra.butCancel fra.seh1 T 1 5 {-st wes} {-t "Відміна" -com "::wndDates res $::dateFrame 0"}}
            {fra.butOk fra.butCancel L 1 5 {-st es} {-t "Вибрати" -com "::wndDates res $::dateFrame 1"}}
        }
        ::wndDates makeWindow $::dateFrame "Проміжок дат"
        ::wndDates paveWindow $::dateFrame $content

        focus $::dateFrame
        grab $::dateFrame

        set res [::wndDates showModal $::dateFrame -focus .dateFrame.fra.dat1]

        destroy $::dateFrame
        ::wndDates destroy       

        if {$res == 1} {
            return [list [clock format [clock scan $::dateBegin -format "%d.%m.%Y"] -format "%Y-%m-%d"] [clock format [clock scan $::dateEnd -format "%d.%m.%Y"] -format "%Y-%m-%d"]]
        }
        return 0
    }

    proc searchHelpAction {} \
    {
        set helpText {
            \m      - початок слова. 
            \M      - кінець слова.
            \y      - початок або кінець слова.
            \Y      - НЕ початок або кінець слова.
            \A      - початок строки.
            \Z      - кінець строки.
            \d      - цифра.
            \w      - символ.
            *       - довільна кількість, або відсутність вираза.
            +       - один або більша кількість.
            ?       - один або відсутність вираза.
            (фа|аф) - або фа, або аф.
            [фа]    - або ф, або а. лише 1 символ.
            [^....] - жоден з символів не входить.
            (.....) - групування виразів.

            \mРО(\w*)[^А]\M  -  слово з 'РО' \mРО + 
                                довільна кількість літер (\w*) + 
                                не 'А' в кінці слова [^А]\M
            
            \m\d+( ?)т[рб]\M  - спочатку цифри
                                далі один або нуль пробілів
                                далі 'т'
                                'р' або 'б' в кінці
        }

        tk_messageBox -message "Пошук по шаблонах" -detail $helpText
        return
    }

    proc fillMenuAction {} \
    {
        set m .mainFrame.menu.file
        $m add command -label "Додати" -command {::pmm::addAction}
        $m add command -label "Редагувати" -command {::pmm::editAction}
        $m add separator
        $m add command -label "Залишки підрозділу по АБ" -command {::pmm::getZalyshkiPidrozdil "АБ"}
        $m add command -label "Залишки підрозділу по ДП" -command {::pmm::getZalyshkiPidrozdil "ДП"}
        $m add command -label "Залишки по АБ" -command {::pmm::getZalyshki "АБ"}
        $m add command -label "Залишки по ДП" -command {::pmm::getZalyshki "ДП"}
        $m add separator
        $m add command -label "Вибір по документу" -command {::pmm::getByDocumentAction}
        $m add command -label "Вибір по ПММ" -command {::pmm::getByPmmAction}
        $m add command -label "Вибір по підрозділу" -command {::pmm::getByPidrozdilAction}
        $m add separator
        $m add command -label "Вихід" -command {::pmm::prog_exit}
        
        set m .mainFrame.menu.view
        $m add command -label "Підрозділи" -command {}
        $m add command -label "Найменування ПММ" -command {}

        set m .mainFrame.menu.sys
        $m add command -label "Допомога по пошуку" -command {::pmm::searchHelpAction}
        $m add separator
        $m add command -label "Зміна паролю" -command {::pmm::changePassAction}
    }

    proc prog_exit {} \
    {
        ::pmm::mysql::close
        exit
    }

    proc changePassAction {} \
    {
        apave::APave create ::chPass
        set ::passFrame .passFrame

        set content {
            {fra  - - - - {-st news -padx 5 -pady 5}}
            {fra.lab1 - - 1 1    {-st es}  {-t "Пароль: "}}
            {fra.ent1 fra.lab1 L 1 9 {-st wes} {-tvar ::new_pass -show {*}}}
            {fra.lab2 fra.lab1 T 1 1 {-st es}  {-t "Повторіть: "}}
            {fra.ent2 fra.lab2 L 1 9 {-st wes} {-tvar ::ver_pass  -show {*}}}
            {fra.seh1 fra.lab2 T 1 10 }
            {fra.butOk fra.seh1 T 1 5 {-st es} {-t "Змінити" -com "::pmm::checkPass"}}
            {fra.butCancel fra.butOk L 1 5 {-st wes} {-t "Відміна" -com "::chPass res $::passFrame 0"}}
        }
        ::chPass makeWindow $::passFrame "Зміна паролю"
        ::chPass paveWindow $::passFrame $content

        focus $::passFrame
        grab $::passFrame

        set res [::chPass showModal $::passFrame -focus .passFrame.fra.fco1]

        destroy $::passFrame
        chPass destroy       
        if {$res == 1} {
            ::pmm::mysql::runQuery "alter user $::db_user identified by \"$::new_pass\" "
            tk_messageBox -message "Пароль змінено!" -detail "Новий пароль - $::new_pass . Не проєбіть."
        }
        return
    }

    proc checkPass {} \
    {
        if {$::new_pass != $::ver_pass} {
            tk_messageBox -message "Паролі не співпадають!" -icon error
        } else {
            ::chPass res $::passFrame 1
        }
        return
    }
    
    proc applyDocument {} \
    {
        if {[lsearch [split $::currentDocument " "] "Виберіть"] != -1} {
            return
        }

        foreach {name date nomer postach} [split $::currentDocument "-"] {
            set ::postach [string trim $postach]
            set ::doc_name [string trim $name]
            set ::doc_nomer [string trim $nomer]
            set ::doc_date [string trim $date]
        }
        focus .addForm.fra.fraValues.entPryhidL
        return
    }
}
