package require sqlite3

sqlite3 db ./doc_catalog

set rows [db eval "select entity_id, doc_date, r_date, date_control from catalog"]

foreach {f1 f2 f3 f4} $rows {
    #set doc_date [clock format [clock scan $f2 -format "%d.%m.%Y"] -format "%Y-%m-%d"]
    #set r_date [clock format [clock scan $f3 -format "%d.%m.%Y"] -format "%Y-%m-%d"]

    if {[string trim $f4] != ""} {
        set date_control [clock format [clock scan $f4 -format "%d.%m.%Y"] -format "%Y-%m-%d"]
        db eval "update catalog set date_control=\"$date_control\" where entity_id=$f1"
    }
#    db eval "update catalog set doc_date=\"$doc_date\", r_date=\"$r_date\" where entity_id=$f1"
}

exit
