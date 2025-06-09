encoding system utf-8

namespace eval ::pmm::winUtils {

    proc centerWindow {w {width 400} {height 250} } {
        after idle "
                update idletasks

                if {[winfo exists $w] == 1} {
                    # centre
                    set xmax \[winfo screenwidth $w\]
                    set ymax \[winfo screenheight $w\]
                    set x \[expr \{(\$xmax - \[winfo reqwidth $w\]) / 2\}\]
                    set y \[expr \{(\$ymax - \[winfo reqheight $w\]) / 2\}\]

                    wm geometry $w \"+\$x+\$y\"
                }
                "
    }

    proc close {wmName} \
    {
        global mainTitle
        destroy $wmName
        wm title . $mainTitle
    }

    proc mapMark {x y can} \
    {
        set ::posX [$can canvasx $x]
        set ::posY [$can canvasy $y]
    }

    proc comboZoom {canv} {
        global mainImage workingImage

        $canv configure -cursor wait
        update

        $canv delete mainMap
        $canv delete mainOverlay
        set width [image width $::mainImage]
        set height [image height $::mainImage]
        set ::workingImage [image create photo wImage]
        set nZoom [string trimright $::cZoom "%"]
        resizephoto $mainImage $workingImage [expr $width*$nZoom/100] [expr $height*$nZoom/100]
        set width [image width $::workingImage]
        set height [image height $::workingImage]
        set scrollregion [list 0 0 $width $height]
        $canv configure -scrollregion $scrollregion
        $canv create image 0 0 -image wImage -anchor nw -tags mainMap
        #$canv create image 0 0 -image overlay -anchor nw -tags mainOverlay

        $canv configure -cursor arrow
        update

        return
    }

    proc initWM {args} {

        # Initializes Tcl/Tk session. Used to be called at the beginning of it.
        #   args - options ("name value" pairs)

        if {!$::apave::_CS_(initWM)} return
        
        lassign [apave::parseOptions $args -cursorwidth $::apave::cursorwidth -theme {clam} \
            -buttonwidth -8 -buttonborder 1 -labelborder 0 -padding 1] \
            cursorwidth theme buttonwidth buttonborder labelborder padding
        set ::apave::_CS_(initWM) 0
        set ::apave::_CS_(CURSORWIDTH) $cursorwidth
        set ::apave::_CS_(LABELBORDER) $labelborder
        #wm withdraw .
        if {$::tcl_platform(platform) eq {windows}} {
            #wm attributes . -alpha 0.0
        }
        # for default theme: only most common settings
        set tfg1 $::apave::_CS_(!FG)
        set tbg1 $::apave::_CS_(!BG)
        if {$theme ne {}} {catch {ttk::style theme use $theme}}
        
        ttk::style map . \
            -selectforeground [list !focus $tfg1 {focus active} $tfg1] \
            -selectbackground [list !focus $tbg1 {focus active} $tbg1]
        ttk::style configure . -selectforeground  $tfg1 -selectbackground $tbg1

        # configure separate widget types
        ttk::style configure TButton -anchor center -width $buttonwidth \
            -relief raised -borderwidth $buttonborder -padding $padding
        ttk::style configure TMenubutton -width 0 -padding 0
        # TLabel's standard style saved for occasional uses
        ttk::style configure TLabelSTD {*}[ttk::style configure TLabel]
        ttk::style configure TLabelSTD -anchor w
        ttk::style map       TLabelSTD {*}[ttk::style map TLabel]
        ttk::style layout    TLabelSTD [ttk::style layout TLabel]
        # ... TLabel new style
        ttk::style configure TLabel -borderwidth $labelborder -padding $padding
        # ... Treeview colors
        set twfg [ttk::style map Treeview -foreground]
        set twfg [apave::putOption selected $tfg1 {*}$twfg]
        set twbg [ttk::style map Treeview -background]
        set twbg [apave::putOption selected $tbg1 {*}$twbg]
        ttk::style map Treeview -foreground $twfg
        ttk::style map Treeview -background $twbg
        # ... TCombobox colors
        ttk::style map TCombobox -fieldforeground [list {active focus} $tfg1 readonly $tfg1 disabled grey]
        ttk::style map TCombobox -fieldbackground [list {active focus} $tbg1 {readonly focus} $tbg1 {readonly !focus} white]

        apave::initPOP .
        apave::initStyles
        apave::initStylesFS name Tahoma size 6
    }
}