#Requires AutoHotkey v2.0

; SSL Renewal Email
    ::renewal ::
    {
        global domain, expirationDate

        clipContent := A_Clipboard

        ; Check if the clipboard content matches the expected format
        if (InStr(clipContent, "Service Ticket #") && InStr(clipContent, " - ") && StrLen(clipContent) >= 10) {
            ; Extract the date (last 10 characters)
            expirationDate := SubStr(clipContent, -10)

            ; Find the positions of the dashes
            firstDashPos := InStr(clipContent, " - ") + 3
            lastDashPos := InStr(clipContent, " - ", , -1)

            ; Extract the domain information between the dashes
            if (firstDashPos > 3 && lastDashPos > firstDashPos) {
                domain := SubStr(clipContent, firstDashPos, lastDashPos - firstDashPos)
            } else {
                domain := "Unable to extract domain information"
            }

            ; Display the extracted information (you can modify this part as needed)
            MsgBox("Domain: " domain "`nExpiration Date: " expirationDate)
        } else {
            MsgBox("Clipboard content does not match the expected format.")
        }

        RecipientName := InputBox("Name of the Recipient", "", "").value
   
        Send
        (
        "Good morning " RecipientName ",
       
        I have the " domain " coming up for renewal on " expirationDate ". Could you verify if this is being renewed internally and on a valid credit card?
        
        Thank you
        
        "
        )
 
        RecipientName := ""
        Domain := ""
        Date := ""
    }    

; Sales Order pasting
    !p::
    {
        ; Store the current active window
        originalWindow := WinGetID("A")

        ; Activate Notepad++
        if WinExist("ahk_exe notepad++.exe")
        {
            WinActivate
            WinWaitActive "ahk_exe notepad++.exe"
            
            ; Move to the end of the document
            Send "^{End}"

            ; Select the last line
            Send "+{Home}"

            ; Copy the selected text
            Send "^c"
            
            SO := "-SO-"

            If (InStr(A_Clipboard, SO) > 0)
            {
                SalesOrder := true
            }
            else
            {
                SalesOrder := false
            }

            while (SalesOrder == false) {
                Send "{Up}"

                Send "{Home}"

                Send "+{End}"

                Send "^c"
                
                Sleep 50

                If (InStr(A_Clipboard, SO) > 0)
                {
                    SalesOrder := true
                }
            }
            
            ; Switch back to the original window
            WinActivate "ahk_id " originalWindow
            WinWaitActive "ahk_id " originalWindow
            
            ; Paste the copied text
            Send "^v"
        }
        else
        {
            MsgBox "Notepad++ is not running."
        }
    }

; Ticket Number collection
    ; Initialize ticket number variable
    global ticketNumber := ""

    ; Use OnClipboardChange instead of overriding Ctrl+C
    OnClipboardChange CheckClipboard

    CheckClipboard(Type) {
        global ticketNumber
        if (Type != 1)  ; Not a text change
            return

        ; Find the position of "Service Ticket #" in the clipboard content
        position := InStr(A_Clipboard, "Service Ticket #")
        
        if (position > 0)
        {
            ; Extract the 7-character ticket number
            ticketNumber := SubStr(A_Clipboard, position + 16, 7)
        }
    }

    ; Alt+3 to paste the ticket number
    !t:: {
        global ticketNumber  ; Ensure we're using the global variable

        if (ticketNumber != "")
        {
            Send ticketNumber
        }
    }

; Minimize Window with back mouse button
    XButton1:: ; Back mouse button minimizes the currently focused window
    {
        WinMinimize "A"
    }

; Email address pasting
    !e:: ; Alt+E to paste my email address
    {
        Send "idarr@expedienttechnology.com"
    }

; Email signature pasting
    !s:: ; Alt+S to paste my email signature
    {
        Send
        (
            "Isaac Darr
            Inside Sales / Solutions Consultant

            Expedient Technology Solutions
            idarr@expedienttechnology.com
            www.expedienttechnology.com

            For Technical Support Contact:
            help@stressfreeit.com or 
            (937) 535-4300"
        )
    }

; Expedient Technology Solutions pasting
    ::ets :: 
    {
        Send "Expedient Technology Solutions"
    }

; Extra clipboard!
    !c:: ; Alt+C to copy to the extra clipboard
    {
        system_clipboard := A_Clipboard
        Send '^c'
        ClipWait
        Sleep 50
        global ahk_clipboard := A_Clipboard
        A_Clipboard := system_clipboard
        system_clipboard:= ""
    }
 
    !v:: ; Alt+V to paste the extra clipboard
    {
        Send "{Text}" ahk_clipboard
    }