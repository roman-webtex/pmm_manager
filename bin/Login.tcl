encoding system utf-8
        
namespace eval ::pmm {

	proc login {} \
	{
    	apave::APave create pave
        pave csSet -2 .

        set loginFrame .loginFrame

        set content {
            {fra  - - - - {-st news -padx 5 -pady 5}}
            {fra.lab1 - - 1 1    {-st es}  {-t "Ім'я: "}}
            {fra.fco1  fra.lab1 L 1 9 {-st we} {-tvar ::db_user -values {@@ $::workingDir/etc/users.ini  @@}}}
            {fra.lab2 fra.lab1 T 1 1 {-st es}  {-t "Пароль: "}}
            {fra.ent2 fra.lab2 L 1 9 {-st wes} {-tvar ::db_password  -show {*}}}
            {fra.seh1 fra.lab2 T 1 10 }
            {fra.butOk fra.seh1 T 1 5 {-st es} {-t "Вхід" -com "::pmm::pave res $loginFrame 1"}}
            {fra.butCancel fra.butOk L 1 5 {-st wes} {-t "Відміна" -com "::pmm::pave res $loginFrame 0"}}
        }
        pave makeWindow $loginFrame "Вхід"
        pave paveWindow $loginFrame $content

        focus $loginFrame
        grab $loginFrame

        set res [pave showModal $loginFrame -focus .loginFrame.fra.fco1]

        destroy $loginFrame
        destroy pave

        if {[string trim $res] == 0} {
            exit
        }

        if {[catch {set ::data [::pmm::mysql::connect]} cerr]} {
            tk_messageBox -type ok -icon error -title "Вхід до програми" -message "Невірний пароль або ім'я користувача"
            tk_messageBox -message $cerr
            exit
        }
        return
    }
}
