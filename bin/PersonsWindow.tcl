encoding system utf-8

namespace eval ::persons {
    
    proc buildMainWindow {} {
        set ::persons_searchVariable ""
        set ::persons_file_name ""
        set ::persons_filePath ""
        set ::persons_iniDir "."
        set ::persons_mainTitle "Дані по персоналу"
        set ::dbpersons_table_name "persons"
        set ::in_vacation 0
        set ::outdate_vacation 0
        set ::all_persons 0
        
        persons::createPopupMenu

        foreach {i icon} {0 add 1 change 2 delete 3 more 4 info 5 double 6 terminal} {
            image create photo PersonsImg$i -data [apave::iconData $icon]
        } 

        set ::attachImg [image create photo tblAttachImg -data [apave::iconData attach]]

        apave::APave create ::persons_pave

        set ::persons_tblcols { 0 { pn } left 0 {} center 0 { ПІБ } left 0 { Тип } left 0 { Початок } center 0 { Закінчення } center 0 { В строю з } center 0 { Примітка } left 0 { День народження } center 0 { ІПН } center 0 { Телефон } center 0 { file_name } center 0 {} center 0 {} center 0 {} center 0 {} center}

        toplevel .personsFrame
        wm title .personsFrame $::persons_mainTitle

        wm geometry .personsFrame "$::window_size"
        
        set content {
            {tool - - - - {pack -side top} {-array { 
            PersonsImg0 {{::persons::addAction} -tooltip "Додати"} sev 5 h_ 1
            PersonsImg1 {{::persons::editAction} -tooltip "Редагувати"} sev 5 h_ 1
            PersonsImg2 {{::persons::outVacationAction} -tooltip "Закрити відпустку"} sev 5 h_ 1
            PersonsImg3 {{::persons::showAction} -tooltip "Інформація по персоналу"} sev 5 h_ 1
            PersonsImg4 {{::persons::showPersonXlsxAction} -tooltip "Інформація по людині"}
            PersonsImg5 {{::persons::showPersonsXlsxAction} -tooltip "Інформація по людях"} sev 5 h_ 1
            PersonsImg6 {{::persons::showScanAction} -tooltip "Скан документу"} sev 5 h_ 1
            ChbAll      { -command {::persons::fillTableAction} -var ::all_persons -t "Всі" -tooltip "Показувати всіх"} sev 5 h_ 1
            ChbVacation { -command {::persons::toggleVacationAction} -var ::in_vacation -t "Відпустки" -tooltip "Люди у відпустках, відрядженнях"}
            ChbOutDate  { -command {::persons::toggleOutdateAction} -var ::outdate_vacation -t "Прострочені" -tooltip "Прострочені відпустки, відрядження"}
            }}}
            {fraPSearch - - - - {pack -side top -fill x -expand 0 -padx 5 -pady 3} {-relief flat}}
            {fraPSearch.entPSearch - - - - {pack -fill x} {-tvar ::persons_searchVariable}}
            {fraPTable - - - - {pack -side top -fill both -expand 1 -padx 5 -pady 3}}
            {fraPTable.TblPMain - - - - {pack -side left -fill both -expand 1} {-h 7 -lbxsel but -columns {$::persons_tblcols}}}
            {fraPTable.sbv fraPTable.tblPMain L - - {pack -side left -after %w}}
            {fraPTable.sbh fraPTable.tblPMain T - - {pack -side left -before %w}}
        }

        ::persons_pave paveWindow .personsFrame $content
        set ::dbPersonsTable [::persons_pave TblPMain]
        $::dbPersonsTable setbusycursor

        $::dbPersonsTable columnconfigure 0 -name por_nom -maxwidth 0 -hide yes
        $::dbPersonsTable columnconfigure 1 -name in_use -maxwidth 1
        $::dbPersonsTable columnconfigure 2 -name pib 
        $::dbPersonsTable columnconfigure 3 -name type -maxwidth 4
        $::dbPersonsTable columnconfigure 4 -name begin_date -maxwidth 8
        $::dbPersonsTable columnconfigure 5 -name end_date -maxwidth 8
        $::dbPersonsTable columnconfigure 6 -name start_date -maxwidth 8
        $::dbPersonsTable columnconfigure 7 -name prymitka -maxwidth 15
        $::dbPersonsTable columnconfigure 8 -name birthday -maxwidth 8
        $::dbPersonsTable columnconfigure 9 -name ipn -maxwidth 10
        $::dbPersonsTable columnconfigure 10 -name phone -maxwidth 25
        $::dbPersonsTable columnconfigure 11 -name f_name -maxwidth 0 -hide yes
        $::dbPersonsTable columnconfigure 12 -name attach -maxwidth 1
        $::dbPersonsTable columnconfigure 13 -name f_type -maxwidth 0 -hide yes 
        $::dbPersonsTable columnconfigure 14 -name s_nom -maxwidth 0 -hide yes 
        $::dbPersonsTable columnconfigure 15 -name online -maxwidth 0 -hide yes 

        $::dbPersonsTable restorecursor
        
        ::persons::fillTableAction 
        
        $::dbPersonsTable setbusycursor

        bind .personsFrame <Escape> {
            set ::persons_inSearch 0
            set ::persons_searchVariable ""
            set ::persons_filePath ""
            wm title .personsFrame $::persons_mainTitle
            ::persons::fillTableAction
        }
        bind .personsFrame <Button-3> {::persons::trackPopupMenu %X %Y}
        bind .personsFrame <Double-Button-1> {::persons::editAction}
        bind .personsFrame <Key-space> {::persons::toggleOblikAction}

        wm protocol .personsFrame WM_DELETE_WINDOW {::persons::prog_exit}

        bind .personsFrame.fraPSearch.entPSearch <Return> {::persons::searchAction}
        focus .personsFrame.fraPSearch.entPSearch

        $::dbPersonsTable restorecursor
    }

    proc fillTableAction {{id_list 0} {where ""}} \
    {
        $::dbPersonsTable setbusycursor

        set addWhere "1"
        set showAll " and in_use = 1 "
        
        if {$::all_persons == 1} {
            set showAll ""
        }

        if {$id_list != 0} {
           set addWhere "entity_id in ($id_list)"
        } elseif {$where != ""} {
           set addWhere $where
        }

        set documentDBRecordSet [::catalog::mysql::getRecordSetWithQuery "*" $::dbpersons_table_name "$addWhere $showAll order by pib"]
        $::dbPersonsTable delete 0 end

        set row 0
        foreach currow $documentDBRecordSet {
          foreach {entity_id pib date_begin date_end date_show in_use prymitka dod_inf ipn phone birthday type file_type attach s_nom online} $currow {

            if {$date_begin != "0000-00-00"} {
                set date_begin [clock format [clock scan $date_begin -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set date_begin ""
            }
            if {$date_end != "0000-00-00"} {
                set date_end [clock format [clock scan $date_end -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set date_end ""
            }
            if {$date_show != "0000-00-00"} {
                set date_show [clock format [clock scan $date_show -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set date_show ""
            }
            if {$birthday != "0000-00-00"} {
                set birthday [clock format [clock scan $birthday -format "%Y-%m-%d"] -format "%d.%m.%Y"]
            } else {
                set birthday ""
            }

            if {$in_use == 1} {
                set in_use "+"
            } else {
                set in_use "-"
            }

            $::dbPersonsTable insert end [list $entity_id $in_use $pib $type $date_begin $date_end $date_show $prymitka $birthday $ipn $phone $dod_inf "" $file_type $s_nom $online]

            if {$date_begin != ""} {
                if {$type == "Відр"} {
                        $::dbPersonsTable rowconfigure $row -bg "#f3ad56"
                } elseif {$type == "Шпит"} {
                        $::dbPersonsTable rowconfigure $row -bg "#b5cbe7"
                } else {
                    if {([clock scan $date_begin -format "%d.%m.%Y"] < [clock seconds]) && ([clock scan $date_end -format "%d.%m.%Y"] > [clock seconds])} {
                        $::dbPersonsTable rowconfigure $row -bg "#f2f21d"
                    } 
                    if {([clock scan $date_show -format "%d.%m.%Y"] < [clock add [clock seconds] 1 days])} {
                        $::dbPersonsTable rowconfigure $row -bg "#e7b5b5"
                    }
                    if {($date_begin == [clock format [clock seconds] -format "%d.%m.%Y"])} {
                        $::dbPersonsTable rowconfigure $row -bg "#a32bc2"
                    }
                }
            }
            
            if {$birthday != "" && $in_use == "+"} {
                if {[clock format [clock scan $birthday -format "%d.%m.%Y"] -format "%d.%m"] == [clock format [clock add [clock seconds] 1 days] -format "%d.%m"]
                 || [clock format [clock scan $birthday -format "%d.%m.%Y"] -format "%d.%m"] == [clock format [clock seconds] -format "%d.%m"]} {
                    $::dbPersonsTable rowconfigure $row -bg "#08bd10"
                    if { $::show_message == 1 } {
                        appendStartupMessageAction $pib $birthday
                    }
                }
            }

            if {$attach == 1} {
                $::dbPersonsTable cellconfigure $row,12 -image $::attachImg
            }

            incr row
          }
        }

        $::dbPersonsTable yview moveto 1.0
        wm title .personsFrame "$::persons_mainTitle ($row записів)"
        focus .personsFrame.fraPSearch.entPSearch
        $::dbPersonsTable restorecursor

        if {$::startup_message != "" && $::show_message == 1} {
            tk_messageBox -message "Вітаємо з Днем Народження!" -detail $::startup_message
            set ::show_message 0
        }

        return
    }

    proc searchAction {} \
    {
        if {[string trim $::persons_searchVariable] == ""} {
            ::persons::fillTableAction
            return
        }

        $::dbPersonsTable setbusycursor

        set searchList [split $::persons_searchVariable ";"]

        set result {}
        set rowCount [$::dbPersonsTable size]

        foreach item [$::dbPersonsTable get 0 $rowCount] {
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
            set ::persons_inSearch 1
            set ::persons_filePath ""
            wm title .personsFrame "Пошук - $::persons_searchVariable"
            $::dbPersonsTable restorecursor
            ::persons::fillTableAction [lrange $result 0 end-1]
        } else {
            tk_messageBox -message "\"$::persons_searchVariable\"" -type ok -detail "По вашому запиту нічого не знайдено" -icon info
            $::dbPersonsTable restorecursor
            ::persons::fillTableAction
        }
        return
    }

    proc editAction {} \
    {
        $::dbPersonsTable setbusycursor
        set row [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0]
        set ::entity_id  [lindex $row 0]
        set ::in_use     [expr {[lindex $row 1] == "+" ? 1 : 0} ]
        set ::pib        [lindex $row 2]
        set ::type       [lindex $row 3]
        set ::date_begin [lindex $row 4]
        set ::date_end   [lindex $row 5]
        set ::date_show  [lindex $row 6]
        set ::prymitka   [lindex $row 7]
        set ::birthday   [lindex $row 8]
        set ::ipn        [lindex $row 9]
        set ::phone      [lindex $row 10]
        set ::file_type  [lindex $row 13]
        if {[lindex $row 11] != ""} {
            set old_file_name [set ::persons_file_name  [lindex $row 11]]
            set old_file_type $::file_type
        } else {
            set old_file_name [set old_file_type [set ::persons_file_name  ""]]
        }
        set attach  [expr {[lindex $row 12] == ""} ? 0 : 1]

        set ::s_nom  [lindex $row 14]
        set ::online  [lindex $row 15]

        $::dbPersonsTable restorecursor
        ::persons::editWindowAction
        $::dbPersonsTable setbusycursor

        if {$::persons_done == 1} {
            if {$::date_begin != ""} {
                set ::date_begin [clock format [clock scan $::date_begin -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::date_begin "0000-00-00"
            }
            if {$::date_end != ""} {
                set ::date_end   [clock format [clock scan $::date_end -format "%d.%m.%Y"] -format "%Y-%m-%d"]
                set ::date_show  [clock format [clock scan $::date_show -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::date_end [set ::date_show "0000-00-00"]
            }

            if {$::persons_file_name != "" && $::persons_file_name != $old_file_name} {
                if {[file exists [file join $::workingDir data scans $::entity_id]] == 0} {
                    file mkdir [file join $::workingDir data scans $::entity_id]
                }
                file copy $::persons_file_name [file join $::workingDir data scans $::entity_id [file tail $::persons_file_name]]
                set ::persons_file_name [file tail $::persons_file_name]
                if {$::file_type == ""} {
                    set ::file_type "Відпускний"
                }
                set attach 1
            }

            if {$::persons_file_name == ""} {
                set ::file_type ""
            }

            if {$::birthday != ""} {
                set ::birthday [clock format [clock scan $::birthday -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::birthday "0000-00-00"
            }

            set query "update $::dbpersons_table_name set pib = \"$::pib\", \
                                          date_begin = \"$::date_begin\", \
                                          date_end = \"$::date_end\", \
                                          date_show = \"$::date_show\", \
                                          prymitka = \"$::prymitka\", \
                                          type = \"$::type\", \
                                          birthday = \"$::birthday\", \
                                          ipn = \"$::ipn\", \
                                          in_use = $::in_use, \
                                          dod_inf = \"$::persons_file_name\", \
                                          dod_type = \"$::file_type\", \
                                          is_attach = $attach , \
                                          s_nom = $::s_nom , \
                                          online = $::online , \
                                          phone = \"$::phone\" where entity_id = $::entity_id"
            ::catalog::mysql::runQuery $query
        }
        set ::persons_file_name ""
        set ::persons_file_name ""
        $::dbPersonsTable restorecursor
        ::persons::searchAction
        return
    }

    proc addAction { {is_copy 0} } \
    {
        $::dbPersonsTable setbusycursor

        if {$is_copy == 0} {
            set old_file_name [set ::file_type [set ::persons_file_name [set ::pib [set ::prymitka [set ::ipn [set ::phone [set ::birthday ""]]]]]]]
            set ::date_begin [set ::date_end [set ::date_show ""]]
            set ::type ""
            set ::in_use 1
            set ::online 1
            set ::s_nom 0
        } else {
            set row [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0]
            set ::in_use     [expr {[lindex $row 1] == "+" ? 1 : 0}]
            set ::pib        [lindex $row 2]
            set ::type       [lindex $row 3]
            set ::prymitka   [lindex $row 7]
            set ::birthday   [lindex $row 8]
            set ::ipn        [lindex $row 9]
            set ::phone      [lindex $row 10]
            set ::file_type  [lindex $row 13]
            set ::s_nom      [lindex $row 14]
            set ::online     [lindex $row 15]
            if {[lindex $row 10] != ""} {
                set old_file_name [set ::persons_file_name  [lindex $row 11]]
            } else {
                set old_file_name [set ::persons_file_name  ""]
            }
        }

        $::dbPersonsTable restorecursor
        ::persons::editWindowAction
        $::dbPersonsTable setbusycursor

        if {$::persons_done == 1} {
            if {$::date_begin != ""} {
                set ::date_begin [clock format [clock scan $::date_begin -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::date_begin "0000-00-00"
            }
            if {$::date_end != ""} {
                set ::date_end   [clock format [clock scan $::date_end -format "%d.%m.%Y"] -format "%Y-%m-%d"]
                set ::date_show  [clock format [clock scan $::date_show -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::date_end [set ::date_show "0000-00-00"]
            }

            if {$::birthday != ""} {
                set ::birthday [clock format [clock scan $::birthday -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set ::birthday "0000-00-00"
            }

            if {$::persons_file_name != ""} {
                file copy $::persons_file_name [file join $::workingDir data [file tail $::persons_file_name]]
                set ::persons_file_name [file tail $::persons_file_name]
                if {$::file_type == ""} {
                    set ::file_type "Відпускний"
                }
                set attach 1
            } else {
                set attach 0
                set ::file_type ""
            }

            set query "insert into $::dbpersons_table_name values (null, \"$::pib\", \"$::date_begin\", \"$::date_end\", \"$::date_show\", $::in_use, \"$::prymitka\", \"$::persons_file_name\", \"$::ipn\", \"$::phone\", \"$::birthday\", \"$::type\", \"$::file_type\", $attach, $::s_nom, $::online )"
            ::catalog::mysql::runQuery $query
        }
        set ::persons_file_name ""
        set ::persons_file_name ""
        $::dbPersonsTable restorecursor
        ::persons::fillTableAction
        return
    }

    proc editWindowAction {} \
    {
        $::dbPersonsTable setbusycursor
        set ::editPersonForm .addPersonForm
        apave::APave create ::editPersonPave
        ::editPersonPave makeWindow .addPersonForm.fra [expr {($::pib eq "") ? "Новий документ" : "Редагування: \[ $::pib \]"}]

        set content {
            {fraInner - - - - {pack -side top -padx 5 -pady 5 -fill both -expand 1}}
            {fraInner.labPib - - 1 1 {-st we} {-t "Прізвище :"}}
            {fraInner.entPib fraInner.labPib L 1 2 {-st we} {-tvar ::pib}}
            {fraInner.labBirthday fraInner.entPib L 1 1 {-st we} {-t "День народження:"}}
            {fraInner.datBirthday fraInner.labBirthday L 1 2 {-st we} {-tvar ::birthday -title {День народження} -dateformat %d.%m.%Y -width 10}}
            {fraInner.labBegin fraInner.labPib T 1 1 {-st we} {-t "З :"}}
            {fraInner.datBegin fraInner.labBegin L 1 1 {-st we} {-tvar ::date_begin -title {Початок} -dateformat %d.%m.%Y -width 10}}
            {fraInner.labEnd fraInner.datBegin L 1 1 {-st we} {-t " по :"}}
            {fraInner.datEnd fraInner.labEnd L 1 1 {-st we} {-tvar ::date_end -title {Завершення} -dateformat %d.%m.%Y -width 10}}
            {fraInner.labShow fraInner.datEnd L 1 1 {-st we} {-t " у строю з :"}}
            {fraInner.datShow fraInner.labShow L 1 1 {-st we} {-tvar ::date_show -title {В строю з} -dateformat %d.%m.%Y -width 10}}
            {fraInner.labType fraInner.labBegin T 1 1 {-st we} {-t "Тип :"}}
            {fraInner.fcoType  fraInner.labType L 1 1 {-st we} {-tvar ::type -values {-list {Відп Відр Шпит БР} }}}
            {fraInner.labPrymitka fraInner.fcoType L 1 1 {-st we} {-t "Примітка :"}}
            {fraInner.entPrymitka fraInner.labPrymitka L 1 3 {-st nwes} {-tvar ::prymitka}}
            {fraInner.labIpn fraInner.labType T 1 1 {-st we} {-t "ІПН :"}}
            {fraInner.entIpn fraInner.labIpn L 1 1 {-st we} {-tvar ::ipn}}
            {fraInner.labPhone fraInner.entIpn L 1 1 {-st we} {-t "Телефон :"}}
            {fraInner.entPhone fraInner.labPhone L 1 3 {-st we} {-tvar ::phone }}
            {fraInner.chbInuse fraInner.labIpn T 1 1 {-st we} {-var ::in_use -t "В строю" }}
            {fraInner.entFileType fraInner.chbInuse L 1 1 {-st we} {-tvar ::file_type }}
            {fraInner.labButton fraInner.entFileType L 1 1 {-st we} {-t "Скан :"}}
            {fraInner.FilName fraInner.labButton L 1 4 {} {-tvar ::persons_file_name -title {Виберіть документ} -filetypes {{{Всі файли} .*}}}}
            {fraInner.chbonline fraInner.chbInuse T 1 1 {-st we} {-var ::online -t "На місті" }}
            {fraInner.entsnom fraInner.chbonline L 1 1 {-st we} {-tvar ::s_nom }}
            {fraButton fraInner T - - {pack -side top -fill x -expand 0}}
            {.butCancel .butSave L 1 1 {pack -side right -padx 5 -pady 3} {-t "Відміна" -com "::editPersonPave res $::editPersonForm 0"}}
            {.butSave .butDelete L 1 1 {pack -side right -padx 5 -pady 3} {-t "Зберегти" -com ::persons::saveCheck}}
            {.butDelete - - 1 1 {pack -side left -padx 5 -pady 3} {-t "Історія" -com ::persons::writeHistory}}
        }
        ::editPersonPave paveWindow .addPersonForm.fra $content 
        wm transient .addPersonForm .personsFrame
        $::dbPersonsTable restorecursor

        set ::persons_done [::editPersonPave showModal $::editPersonForm -focus .addPersonForm.fra.fraInner.entPib]
        destroy $::editPersonForm
        ::editPersonPave destroy
        return
    }

    proc outVacationAction {} \
    {
        if {[lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 4] == "" || [$::dbPersonsTable curselection] == ""} {
            return
        }

        if {[tk_messageBox -type yesno -icon question -message "Закрити відпустку (відрядження) [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 2] ?"] == yes} {
            set entity_id [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0]
            set date_begin [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 4]
            set date_end  [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 5]
            set type [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 3]
            set prymitka [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 7]
            set dod_inf [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 11]
            set file_type [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 13]

            set date_begin [clock format [clock scan $date_begin -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            set date_end   [clock format [clock scan $date_end -format "%d.%m.%Y"] -format "%Y-%m-%d"]

            set query "insert into persons_history values (null, $entity_id, \"$date_begin\", \"$date_end\", \"$type\", \"$prymitka\", \"$dod_inf\", \"$file_type\")"
            ::catalog::mysql::runQuery $query
            set query "update $::dbpersons_table_name set date_begin='0000-00-00', date_end='0000-00-00', date_show='0000-00-00', type='', prymitka='', dod_inf='', dod_type='' where entity_id = $entity_id"
            ::catalog::mysql::runQuery $query
        }
        ::persons::fillTableAction
        return
    }

    proc saveCheck {} \
    {
        set detail_message ""
        if {$::pib == ""} {
            append detail_message "Не заповнено прізвище\n"
        }
        if {$::date_begin != "" && $::type == ""} {
            append detail_message "Не заповнений тип\n"
        }
        if {($::date_begin != "" && $::date_end == "" && $::type != "Шпит") || ($::date_begin == "" && $::date_end != "") || ($::date_begin != "" && $::date_end != "" && $::date_show == "")} {
            append detail_message "Невірні дати\n"
        }
        if {$detail_message == ""} {
            ::editPersonPave res $::editPersonForm 1
        } else {
            tk_messageBox -message "Запис містить помилки!" -detail $detail_message -icon error
        }
    }

    proc writeHistory {} \
    {
        if {[tk_messageBox -type yesno -icon question -message "Перенести запис до історії?"] == yes} {
            if {$::date_begin != ""} {
                set date_begin [clock format [clock scan $::date_begin -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set date_begin "0000-00-00"
            }
            
            if {$::date_end != ""} {
                set date_end   [clock format [clock scan $::date_end -format "%d.%m.%Y"] -format "%Y-%m-%d"]
            } else {
                set date_end "0000-00-00"
            }

            set query "insert into persons_history values (null, $::entity_id, \"$date_begin\", \"$date_end\", \"$::type\", \"$::prymitka\", \"$::persons_file_name\", \"$::file_type\")"
            ::catalog::mysql::runQuery $query
            set ::date_begin ""
            set ::date_end   ""
            set ::date_show  ""
            set ::type       ""
            set ::prymitka   ""
            set ::persons_file_name  ""
            set ::file_type  ""
        }
        return
    }

    proc showScanAction {} \
    {
        if {[$::dbPersonsTable curselection] == ""} {
            return
        }
        set file_name [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 11]
        set file_type [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 13]

        if {[winfo exists .popupMenuScans] != 0} {
            destroy .popupMenuScans
        }
        set popupMenuScans [menu .popupMenuScans]
        
        set historyRecordSet [::catalog::mysql::getRecordSetWithQuery "dod_inf, dod_type" "persons_history" "person_id = [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0]"]
        foreach row $historyRecordSet {
            if {[set f_name [lindex $row 0]] != "" && [file exists [file nativename [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $f_name]]] } {
                $popupMenuScans add command -label "[lindex $row 1] від [clock format [file mtime [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $f_name]] -format "%d.%m.%Y"]" \
                    -command [list ::persons::showAttachAction [file nativename [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $f_name]]]
            }
        }

        if {$file_name != ""  && [file exists [file nativename [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $file_name]]]} {
            $popupMenuScans add command -label "$file_type від [clock format [file mtime [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $file_name]] -format "%d.%m.%Y"]" \
                -command [list ::persons::showAttachAction [file nativename [file join $::workingDir data scans [lindex [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0] 0] $file_name]]]
        }

        tk_popup $popupMenuScans [expr {([winfo screenwidth .] / 2) - [winfo reqwidth $popupMenuScans]}] 200

        return
    }

    proc showAttachAction {filename} \
    {
        catch {[exec $::explorer $filename] errorMessage }
    }

    proc toggleVacationAction {} \
    {
        if {$::in_vacation == 1} {
            set ::outdate_vacation 0;
            ::persons::fillTableAction 0 "date_begin != '0000-00-00' and date_begin <= '[clock format [clock seconds] -format "%Y-%m-%d"]' and date_end >= '[clock format [clock seconds] -format "%Y-%m-%d"]'"
        } else {
            ::persons::fillTableAction 
        }
        return
    }

    proc toggleOutdateAction {} \
    {
        if {$::outdate_vacation == 1} {
            set ::in_vacation 0;
            ::persons::fillTableAction 0 "date_begin != '0000-00-00' and date_end <= '[clock format [clock seconds] -format "%Y-%m-%d"]'"
        } else {
            ::persons::fillTableAction 
        }
        return
    }
    

    proc showAction {} \
    {
        set allPersons [::catalog::mysql::runQuery "select count(*) from $::dbpersons_table_name where in_use = 1"]
        set onPlace    [::catalog::mysql::runQuery "select count(*) from $::dbpersons_table_name where in_use = 1 and date_begin = '0000-00-00' or date_begin >= '[clock format [clock seconds] -format "%Y-%m-%d"]'"]
        set onVacation [::catalog::mysql::runQuery "select count(*) from $::dbpersons_table_name where in_use = 1 and date_begin != '0000-00-00' and type = 'Відп' and date_begin <= '[clock format [clock seconds] -format "%Y-%m-%d"]'"]
        set onCommand  [::catalog::mysql::runQuery "select count(*) from $::dbpersons_table_name where in_use = 1 and date_begin != '0000-00-00' and type = 'Відр' and date_begin <= '[clock format [clock seconds] -format "%Y-%m-%d"]'"]
        set onHospital [::catalog::mysql::runQuery "select count(*) from $::dbpersons_table_name where in_use = 1 and date_begin != '0000-00-00' and type = 'Шпит' and date_begin <= '[clock format [clock seconds] -format "%Y-%m-%d"]'"]
        set outPlaceRecordSet [::catalog::mysql::getRecordSetWithQuery "pib, date_begin, date_end, date_show, type" $::dbpersons_table_name "in_use = 1 and date_begin != '0000-00-00' and date_begin <= '[clock format [clock seconds] -format "%Y-%m-%d"]' order by date_begin"]   
        
        ::mkxlsx::xlsxInit

        ::mkxlsx::addCell A 1 "Всього" $::mkxlsx::textStyle
        ::mkxlsx::addCell B 1 $allPersons $::mkxlsx::headerStyle

        ::mkxlsx::addCell A 2 "В строю" $::mkxlsx::textStyle
        ::mkxlsx::addCell B 2 $onPlace $::mkxlsx::headerStyle

        ::mkxlsx::addCell A 3 "Відпустка" $::mkxlsx::textStyle
        ::mkxlsx::addCell B 3 $onVacation $::mkxlsx::headerStyle

        ::mkxlsx::addCell A 4 "Відрядження" $::mkxlsx::textStyle
        ::mkxlsx::addCell B 4 $onCommand $::mkxlsx::headerStyle

        ::mkxlsx::addCell A 5 "Шпиталь" $::mkxlsx::textStyle
        ::mkxlsx::addCell B 5 $onHospital $::mkxlsx::headerStyle


        ::mkxlsx::merge "A7:E7"
        ::mkxlsx::addCell A 7 "Дані по відсутніх" $::mkxlsx::boldStyle

        ::mkxlsx::addCell A 8 "Прізвище" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell B 8 "Тип" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell C 8 "З" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell D 8 "По" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell E 8 "З'явитися" $::mkxlsx::headerFilledStyle

        set commandStyle [::mkxlsx::createStyle fillId [::mkxlsx::createFill patternType "solid" bgColor "FFFF0000" fgColor "99AA99AA"] vartical "top" horizontal "center"]
        set outdateStyle [::mkxlsx::createStyle fillId [::mkxlsx::createFill patternType "solid" bgColor "FFFF0000" fgColor "FFFF0000"] vertical "top" horizontal "center"]

        set rn 9
        foreach row $outPlaceRecordSet {
            foreach {pib date_begin date_end date_show type} $row {
                ::mkxlsx::addCell A $rn $pib $::mkxlsx::textStyle
                if {$type == "Відр"} {
                    ::mkxlsx::addCell B $rn $type. $commandStyle
                } else {
                    ::mkxlsx::addCell B $rn $type. $::mkxlsx::dateStyle
                }
                ::mkxlsx::addCell C $rn [clock format [clock scan $date_begin -format {%Y-%m-%d}] -format {%d.%m.%Y}] $::mkxlsx::dateStyle
                if {$date_end != "0000-00-00"} {
                    if {([clock scan $date_show -format "%Y-%m-%d"] < [clock add [clock seconds] 1 days])} {
                        ::mkxlsx::addCell D $rn [clock format [clock scan $date_end -format {%Y-%m-%d}] -format {%d.%m.%Y}] $outdateStyle
                    } else {
                        ::mkxlsx::addCell D $rn [clock format [clock scan $date_end -format {%Y-%m-%d}] -format {%d.%m.%Y}] $::mkxlsx::dateStyle
                    }
                } else {
                    ::mkxlsx::addCell D $rn " " $::mkxlsx::dateStyle
                }
                if {$date_show != "0000-00-00"} {
                    ::mkxlsx::addCell E $rn [clock format [clock scan $date_show -format {%Y-%m-%d}] -format {%d.%m.%Y}] $::mkxlsx::dateStyle
                } else {
                    ::mkxlsx::addCell E $rn " " $::mkxlsx::dateStyle
                }
                incr rn
            }
        }
        ::mkxlsx::setColWidth 1 35
        ::mkxlsx::setColWidth 4
        ::mkxlsx::writeXml
        ::mkxlsx::mkzip [file nativename [file join $::workingDir out user_data.xlsx]] -directory [file nativename [file join $::workingDir etc xlsx user_data]]
        catch { [exec $::explorer [file nativename [file join $::workingDir out user_data.xlsx]]] errorMessage }
        return
    }

    proc createPopupMenu {} \
    {
        set ::popMenuPersons [menu .popupMenuPersons]
        $::popMenuPersons add command -label "Редагувати" -command {::persons::editAction}
        $::popMenuPersons add command -label "Додатки" -command {::persons::showScanAction}
        $::popMenuPersons add separator
        $::popMenuPersons add command -label "Зняти з обліку / Поставити на облік" -command {::persons::toggleOblikAction}
    }

    proc trackPopupMenu {x y} \
    {
        if {[$::dbPersonsTable curselection] == ""} {
            return
        }
        tk_popup $::popMenuPersons $x $y
        return
    }

    proc toggleOblikAction {} \
    {
        set row [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0]
        set entity_id [lindex $row 0]
        set in_use [expr {[lindex $row 1] == "+" ? 0 : 1}]
        ::catalog::mysql::runQuery "update $::dbpersons_table_name set in_use = $in_use where entity_id = $entity_id"
        ::persons::fillTableAction
        return
    }

    proc appendStartupMessageAction {pib birthday} \
    {
        set dat [clock format [clock scan $birthday -format "%d.%m.%Y"] -format "%d.%m"]
        set pib [split $pib " "]
        set nam [lindex $pib 1]
        set bat [lindex $pib 2]
        if {[string range $nam [string length $nam]-1 [string length $nam]] == "а"} {
            set nam [string replace $nam [string length $nam]-1 [string length $nam] "і"]
        } elseif {[string range $nam [string length $nam]-1 [string length $nam]] == "я"} {
            set nam [string replace $nam [string length $nam]-1 [string length $nam] "ї"]
        } elseif {[string range $nam [string length $nam]-1 [string length $nam]] == "й"} {
            set nam [string replace $nam [string length $nam]-1 [string length $nam] "ю"]
        } else {
            set nam $namу
        }
        if {[string range $bat [string length $bat]-1 [string length $bat]] == "а"} {
            set bat [string replace $bat [string length $bat]-1 [string length $bat] "і"]
        } else {
            set bat $batу
        }
        set cou [expr {[clock format [clock seconds] -format "%Y"] - [clock format [clock scan $birthday -format "%d.%m.%Y"] -format "%Y"]}]
        if {[expr {$cou % 10}] == 1} {
            set t "рік"
        } elseif {[expr {$cou % 10}] > 1 && [expr {$cou % 10}] < 5} {
            set t "рокі"
        } else {
            set t "років"
        }
        append ::startup_message "$nam $bat $dat - $cou $t!\n"
        return
    }

    proc prog_exit {args} \
    {
        destroy .personsFrame
        destroy .popupMenuPersons
        ::persons_pave destroy
    }

    proc showPersonXlsxAction {} \
    {
        if {[$::dbPersonsTable curselection] == ""} {
            tk_messageBox -detail "Потрібно вибрати людину"
            return
        }
        ::mkxlsx::xlsxInit

        set row [lindex [$::dbPersonsTable get [$::dbPersonsTable curselection] [$::dbPersonsTable curselection]] 0]
        set entity_id  [lindex $row 0]
        set pib        [lindex $row 2]
        set type       [lindex $row 3] 
        set date_begin [lindex $row 4]
        set date_end   [lindex $row 5]
        set date_show  [lindex $row 6]
        set prymitka   [lindex $row 7]
        set historyRecordSet [::catalog::mysql::getRecordSetWithQuery "*" "persons_history" "person_id = $entity_id order by date_begin desc"]

        ::mkxlsx::addCell A 1 $pib $::mkxlsx::boldStyle
        ::mkxlsx::merge "A1:D1"

        ::mkxlsx::addCell A 2 "Тип" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell B 2 "З" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell C 2 "По" $::mkxlsx::headerFilledStyle
        ::mkxlsx::addCell D 2 "Примітка" $::mkxlsx::headerFilledStyle
        
        if {$type == "Шпит"} {
            ::mkxlsx::addCell A 3 $type. $::mkxlsx::textStyle
            ::mkxlsx::addCell B 3 [clock format [clock scan $date_begin -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
            ::mkxlsx::addCell C 3 " " $::mkxlsx::dateStyle
            ::mkxlsx::addCell D 3 "н.п. $prymitka" $::mkxlsx::textStyle
        } elseif {$date_begin != "" && ([clock scan $date_begin -format "%d.%m.%Y"] < [clock seconds])} {
            ::mkxlsx::addCell A 3 $type. $::mkxlsx::textStyle
            ::mkxlsx::addCell B 3 [clock format [clock scan $date_begin -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
            if {$date_end != ""} {
                ::mkxlsx::addCell C 3 [clock format [clock scan $date_end -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                ::mkxlsx::addCell D 3 "в строю з [clock format [clock scan $date_show -format "%d.%m.%Y"] -format "%d.%m.%Y"]" $::mkxlsx::textStyle
            } else {
                ::mkxlsx::addCell C 3 " " $::mkxlsx::dateStyle
                ::mkxlsx::addCell D 3 " " $::mkxlsx::dateStyle
            }
        }
        set rn 4        
        foreach row $historyRecordSet {
            foreach {nom rel date_begin date_end type prymitka dod_inf dod_type} $row {
                if {$type == "Шпит"} {
                    set addon "н.п. "
                } else {
                    set addon ""
                }
                ::mkxlsx::addCell A $rn $type. $::mkxlsx::textStyle
                ::mkxlsx::addCell B $rn [clock format [clock scan $date_begin -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                if {$date_end != "0000-00-00"} {
                    ::mkxlsx::addCell C $rn [clock format [clock scan $date_end -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                } else {
                    ::mkxlsx::addCell C $rn " " $::mkxlsx::dateStyle
                }
                ::mkxlsx::addCell D $rn "$addon$prymitka" $::mkxlsx::textStyle
                incr rn
            }
        }
        ::mkxlsx::setColWidth 4 45
        ::mkxlsx::writeXml
        ::mkxlsx::mkzip [file nativename [file join $::workingDir out user_data.xlsx]] -directory [file nativename [file join $::workingDir etc xlsx user_data]]
        catch { [exec $::explorer [file nativename [file join $::workingDir out user_data.xlsx]]] errorMessage }
        return        
    }

    proc showPersonsXlsxAction {} \
    {
        ::mkxlsx::xlsxInit

        set rn 1

        for {set i 0} {$i<[$::dbPersonsTable size]} {incr i} {
            set row [$::dbPersonsTable get $i]
            set entity_id  [lindex $row 0]
            set pib        [lindex $row 2]
            set type       [lindex $row 3] 
            set date_begin [lindex $row 4]
            set date_end   [lindex $row 5]
            set date_show  [lindex $row 6]
            set prymitka   [lindex $row 7]
            set historyRecordSet [::catalog::mysql::getRecordSetWithQuery "*" "persons_history" "person_id = $entity_id order by date_begin desc"]

            if {$historyRecordSet == {} && ($date_begin == "" || [clock scan $date_begin -format "%d.%m.%Y"] > [clock seconds])} {
                continue
            }

            ::mkxlsx::addCell A $rn $pib $::mkxlsx::boldStyle
            ::mkxlsx::merge "A$rn:D$rn"
            incr rn
 
            ::mkxlsx::addCell A $rn "Тип" $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell B $rn "З" $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell C $rn "По" $::mkxlsx::headerFilledStyle
            ::mkxlsx::addCell D $rn "Примітка" $::mkxlsx::headerFilledStyle
            
            if {[string trim $type] == "Шпит"} {
                incr rn 
                ::mkxlsx::addCell A $rn $type. $::mkxlsx::textStyle
                ::mkxlsx::addCell B $rn [clock format [clock scan $date_begin -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                ::mkxlsx::addCell C $rn " " $::mkxlsx::dateStyle
                ::mkxlsx::addCell D $rn "н.п. $prymitka" $::mkxlsx::textStyle
            } elseif {$date_begin != "" && ([clock scan $date_begin -format "%d.%m.%Y"] < [clock seconds])} {
                incr rn 
                ::mkxlsx::addCell A $rn $type. $::mkxlsx::textStyle
                ::mkxlsx::addCell B $rn [clock format [clock scan $date_begin -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                if {$date_end != ""} {
                    ::mkxlsx::addCell C $rn [clock format [clock scan $date_end -format "%d.%m.%Y"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                    ::mkxlsx::addCell D $rn "в строю з [clock format [clock scan $date_show -format "%d.%m.%Y"] -format "%d.%m.%Y"]" $::mkxlsx::textStyle
                } else {
                    ::mkxlsx::addCell C $rn " " $::mkxlsx::dateStyle
                    ::mkxlsx::addCell D $rn " " $::mkxlsx::textStyle
                }
            }
        
            foreach row $historyRecordSet {
                foreach {nom rel date_begin date_end type prymitka dod_inf dod_type} $row {
                    incr rn
                    if {$type == "Шпит"} {
                        set addon "н.п. "
                    } else {
                        set addon ""
                    }
                    ::mkxlsx::addCell A $rn $type. $::mkxlsx::textStyle
                    ::mkxlsx::addCell B $rn [clock format [clock scan $date_begin -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                    if {$date_end != "0000-00-00"} {
                        ::mkxlsx::addCell C $rn [clock format [clock scan $date_end -format "%Y-%m-%d"] -format "%d.%m.%Y"] $::mkxlsx::dateStyle
                    } else {
                        ::mkxlsx::addCell C $rn " " $::mkxlsx::dateStyle
                    }
                    ::mkxlsx::addCell D $rn "$addon$prymitka" $::mkxlsx::textStyle
                }
            }
            incr rn 2
        }
        ::mkxlsx::setColWidth 4 45
        ::mkxlsx::writeXml
        ::mkxlsx::mkzip [file nativename [file join $::workingDir out user_data.xlsx]] -directory [file nativename [file join $::workingDir etc xlsx user_data]]
        catch { [exec $::explorer [file nativename [file join $::workingDir out user_data.xlsx]]] errorMessage }
        return        
    }
}
