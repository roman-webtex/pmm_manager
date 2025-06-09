encoding system utf-8

namespace eval ::pmm::sysutils {

	proc getDate {variable count interval} \
	{
		return [clock format [clock add [clock scan $variable -format {%d.%m.%Y}] $count $interval] -format {%d.%m.%Y}]
	}

	proc toEng {str} \
	{
		string map -nocase { ё jo й j ц c у u к k е e н n г g ш sh щ sh з z х h ї i ф f і i в v а a п p р r о o л l д d ж zh є je я ja ч ch с s м m и y т t ь _ б b ю yu " " _ } $str
	}

        proc getCol {num} {
            if {$num <= 26} {
                return [::pmm::sysutils::numMap $num]
            } else {
                set first [expr {$num / 26 }]
                set second [expr {$num % 26}]
                if {$second == 0} {
                    set first [expr {$first - 1}]
                }
                return [::pmm::sysutils::numMap $first][::pmm::sysutils::numMap $second]
            }
        }
        
        proc numMap {num} {
            return [string map {20 T 21 U 22 V 23 W 24 X 25 Y 26 Z 10 J 11 K 12 L 13 M 14 N 15 O 16 P 17 Q 18 R 19 S 1 A 2 B 3 C 4 D 5 E 6 F 7 G 8 H 9 I 0 Z} $num]
        }

	proc showProgress {{show 0} {currVal 0} {maxVal 100} {message ""}} \
	{
		if {[winfo exist .progressWindow] == 0} {
			set editForm .progressWindow
			toplevel $editForm
			wm title $editForm "Зачекайте"
			wm attributes $editForm -topmost 1
			::pmm::winUtils::centerWindow $editForm
			focus $editForm
			grab $editForm

    		ttk::frame $editForm.innerFrame
			pack $editForm.innerFrame -fill both -expand true
			grid [ttk::label $editForm.innerFrame.progressMessage -text $message]
			grid $editForm.innerFrame.progressMessage -sticky w -padx 3 -pady 2
			grid [ttk::progressbar $editForm.innerFrame.progress -maximum $maxVal -mode determinate -length 400]
			grid $editForm.innerFrame.progress -sticky w -padx 3 -pady 10
		}

		if {[winfo exist .progressWindow] && $show == 0} {
			destroy .progressWindow
		} elseif {[winfo exist .progressWindow]} {
			.progressWindow.innerFrame.progress configure -value $currVal -maximum $maxVal
			.progressWindow.innerFrame.progressMessage configure -text $message
		}
		update
	}

	proc showInfo {{text ""}} \
	{
		.bottomFrame.infoLabel configure -text $text
	}

	proc regExps {} \
	{
		set rDate "\m(([0-2][1-9]|[1-3][01])[-./](0[1-9]|1[0-2])|(0[1-9]|1[0-2])[-./]([0-2][1-9]|[1-3][01]))[-./]([0-9]{1,4})\M"
		set rEuDate "\m(([0-2][1-9]|[1-3][01])[-./](0[1-9]|1[0-2])[-./]([0-9]{1,4}))\M"
		set rUsDate "\m((0[1-9]|1[0-2])[-./]([0-2][1-9]|[1-3][01])[-./]([0-9]{1,4}))\M"
	}

	proc reduceLine {line width {side "right"}} \
	{
		if {[string length $line] > $width} {
			if {$side eq "right"} {
				set line [string range $line 0 $width]...
			} else {
				set line [string range $line 0 10]...[string range $line [expr {[string length $line] - $width + 13}] end]
			}
		}
		return $line
	}

	proc waitcur {{isShow 1}} \
	{
		if {$isShow == 1} {
			. configure -cursor watch
		} else {
			. configure -cursor arrow
		}
	}
}