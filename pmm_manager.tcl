encoding system utf-8

package require Tk
package require tdbc
package require tablelist
package require BWidget
package require apave
package require tdom
package require uuid
namespace import msgcat::mc

set tcl_precision 12
set ::data 0
set ::mainTitle ""
set ::workingDir [file dirname [file normalize [info script]]]
set ::table_year [clock format [clock seconds] -format "%Y"]
set ::yearList {"{ 2024 }" "{ 2025 }"}
set ::show_message 1
set ::startup_message ""
set ::fSize 8

foreach filename [glob -nocomplain [file join $::workingDir bin *.tcl]] {
    source $filename
}

package require tdbc::$::db_engine

switch $::tcl_platform(platform) {     
    windows {         
        set ::filename_prefix $::windows_prefix     
        set ::explorer explorer
        set ::window_size "[winfo screenwidth .]x[expr {[winfo screenheight .] - 100}]+0+0"
    } 
    unix {         
        set ::filename_prefix $::unix_prefix     
        set ::explorer xdg-open
        set ::window_size "[winfo screenwidth .]x[winfo screenheight .]+0+0"
    } 
}

namespace eval ::pmm {
}

foreach font_name [font names] {
    font configure $font_name -size $::fSize
}
    
proc ::main {} {
    apave::initWM
    ::pmm::login
    ::pmm::buildMainWindow
}

::main

