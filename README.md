# vbscript

'Variables available on this screen can be declared and initialized here.


'This procedure is executed just once when this screen is open.
Sub Screen_OnOpen()

$TempIndexComboBox = $IndexComboBox

End Sub

'This procedure is executed continuously while this screen is open.
Sub Screen_WhileOpen()

If $TempIndexComboBox <> $IndexComboBox Then

    Select Case $IndexComboBox
        Case 0
            $Open("032C1")
        Case 1
            $Open("032C2")
        Case 2
            $Open("032C3")
        Case 3
            $Open("032C4")
        Case 4
            $Open("032C5")
        Case 5
            $Open("032C6")
        Case 6
            $Open("032C7")
        Case 7
            $Open("032C8")
        Case 8
            $Open("032C9")
        Case 9
            $Open("032C10")
        Case 10
            $Open("032C11")
        Case 11
            $Open("032C12")
        Case 12
            $Open("032C13")
        Case 13
            $Open("032C14")
        Case 14
            $Open("032C15")
        Case 15
            $Open("032C16")
        Case 16
            $Open("032C17")
        Case 17
            $Open("032C18")
        Case 18
            $Open("032C19")
        Case 19
            $Open("032C20")
        Case 20
            $Open("032C21")
        Case 21
            $Open("032C22")
        Case 22
            $Open("032C23")
        Case 23
            $Open("032C24")
        Case 24
            $Open("032C25")
        Case 25
            $Open("032C26")
        Case 26
            $Open("032C27")
        Case 27
            $Open("032C28")
        Case 28
            $Open("032C29")
        Case 29
            $Open("032C30")
        Case 30
            $Open("032C31")
        Case 31
            $Open("032C32")
        Case 32
            $Open("032C33")
        Case 33
            $Open("032C34")
        Case 34
            $Open("032C35")
        Case 35
            $Open("032C36")
        Case 36
            $Open("032C37")
        Case 37
            $Open("032C38")
        Case 38
            $Open("032C39")
        Case 39
            $Open("032C40")
        Case 40
            $Open("032C41")
        Case 41
            $Open("032C42")
        Case 42
            $Open("032C43")
        Case 43
            $Open("032C44")
        Case 44
            $Open("032C45")
    End Select

    'update temp
    $TempIndexComboBox = $IndexComboBox

End If

End Sub

'This procedure is executed just once when this screen is closed.
Sub Screen_OnClose()

End Sub
