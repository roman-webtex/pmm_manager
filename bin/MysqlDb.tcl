encoding system utf-8

namespace eval ::pmm::mysql {

    proc connect {} \
    {
        ::tdbc::mysql::connection create ::data -user $::db_user -password $::db_password -host $::db_host -port $::db_port -db $::db_name 
        #-interactive true
    }
    
    proc close {} \
    {
        ::data close
    }

    proc getRecordSetWithQuery { what from where } \
    {
        set query "select $what from $from where $where"
#        tk_messageBox -message $query
        return [::pmm::mysql::runQuery $query]
    }

    proc getDistinctRecordSetWithQuery { what from where } \
    {
        set query "select distinct $what from $from where $where"
        return [::pmm::mysql::runQuery $query]
    }

    proc runQuery {query} \
    {
        if {$::data == 0} {
            ::pmm::mysql::connect
        }
        set stm [::data prepare $query]
        set result [$stm execute]
        set retval [$result allrows -as lists]
        $result close
        $stm close
        return $retval
    }
    
    proc runSelect {query} \
    {
        return [::pmm::mysql::runQuery $query]
    }

}