encoding system utf-8

namespace eval ::catalog::sqlite {
    proc createTable {table_name query} \
    {
        sqlite3 zsu-db $::dbName
        set createQuery "create table $table_name ($query)"
        zsu-db eval $createQuery
        zsu-db close
    }

    proc getRecordSet { table_name } \
    {
        sqlite3 zsu-db $::dbName
        set query "select * from $table_name"
        set retval [zsu-db eval $query]
        zsu-db close
        return $retval
    }

    proc getRecordSetWithQuery { what from where } \
    {
        sqlite3 zsu-db $::dbName
        set query "select $what from $from where $where"
        set retval [zsu-db eval $query]
        zsu-db close
        return $retval
    }

    proc getRecordSetWithJoin { what from join } \
    {
        sqlite3 zsu-db $::dbName
        set query "select $what from $from $join"
        set retval [zsu-db eval $query]
        zsu-db close
        return $retval
    }

    proc load {table_name where} \
    {
        sqlite3 zsu-db $::dbName
        set query "select * from $table_name where $where"
        set retval [zsu-db eval $query]
        zsu-db close
        return $retval
    }

    proc saveRecord {table_name update where} \
    {
        sqlite3 zsu-db $::dbName
        set query "update $table_name set $update where $where"
        zsu-db eval $query
        zsu-db close
   }

    proc deleteRecord {table_name where} \
    {
        sqlite3 zsu-db $::dbName
        set query "delete from $table_name where $where"
        zsu-db eval $query
        zsu-db close
   }

   proc addNew {table_name values} \
   {
        foreach val $values {
            if {$val != "NULL"} {
                lappend query "'$val'"
            } else {
                lappend query "$val"
            }
            lappend query ","    
        }
        set query "insert into $table_name values ([lrange $query 0 end-1])"
       
        sqlite3 zsu-db $::dbName
        zsu-db eval $query
        zsu-db close
   }

   proc runQuery {query} \
   {
        sqlite3 zsu-db $::dbName
        set result [zsu-db eval $query]
        zsu-db close
        return $result       
   }

}