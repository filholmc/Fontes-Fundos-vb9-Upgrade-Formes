 Option Strict      On
 Option Explicit    On

 Friend Class           formAtuVer
        Inherits        System.Windows.Forms.Form

Private fbooArqExe  As  Boolean

Private fdatVerLoc  As  Date
Private fdatVerRde  As  Date

Private fstrArgmto  As  String()             =     Environment.GetCommandLineArgs
Private Sub  fcmbF03Atu_Click               (ByVal Objt       As  Object, ByVal Args As System.                       EventArgs) _
Handles      fcmbF03Atu.Click
             rlsLimparMensagens
             flblMsg002.Text                 =    "Copiando "
             flblMsg002.Refresh
             flblMsg004.Text                 =     fstrArgmto(3)
             flblMsg004.Refresh
             flblMsg006.Text                 =    "para"
             flblMsg006.Refresh
             flblMsg008.Text                 =     fstrArgmto(2) & " aguarde..."
             flblMsg008.Refresh

        If   fbooArqExe
             My.Computer.FileSystem.CopyDirectory (fstrArgmto(3).Substring(0, fstrArgmto(3).LastIndexOf("\")),                   _
                                                   fstrArgmto(2).Substring(0, fstrArgmto(2).LastIndexOf("\")), True)
        Else
             My.Computer.FileSystem.CopyFile      (fstrArgmto(3),             fstrArgmto(2),                   True)
        End  If

             MsgBox                               ("Atualização concluída com sucesso!",    MsgBoxStyle.Information)

             rlsLimparMensagens

        If   fbooArqExe
             flblMsg004.Text                 =    "Reexecutando o Módulo "            &     fstrArgmto(1) & "..."
             flblMsg004.Refresh

             Shell(fstrArgmto(2),                  AppWinStyle.NormalFocus)
        End  If
        Me . Close
End     Sub
Private Sub  formAtuVer_KeyDown             (ByVal Objt       As  Object, ByVal Args As                            KeyEventArgs) _
Handles              Me.KeyDown
Select  Case "F"      & Format              (Args. KeyCode - 111, "0")
        Case "F3"
        If   fcmbF03Atu.Enabled
             fcmbF03Atu_Click               (Objt, Args)
        End  If
End     Select
End     Sub
Private Sub  formAtuVer_Load                (ByVal Objt       As  Object, ByVal Args As System.                       EventArgs) _
Handles              Me.Load
        If   fstrArgmto.Length               >     1
             fbooArqExe                      =    "exe"       =   fstrArgmto(2).Substring   (fstrArgmto(2).Length - 3,   3)
             fdatVerLoc                      =     FileDateTime  (fstrArgmto(2))
             fdatVerRde                      =     FileDateTime  (fstrArgmto(3))

             flblMsg001.Text                 =    "Atualização do Módulo "      &            fstrArgmto(1)            & "."
             flblMsg001.Refresh
             flblMsg003.Text                 =    "Há uma Versão mais recente " &    CStr(If(fbooArqExe,  "", "da Ajuda "))    & _
                                                   "deste Módulo em"
             flblMsg003.Refresh
             flblMsg004.Text                 =                    fstrArgmto(3).Substring(0, fstrArgmto(3).LastIndexOf("\") +  1)
             flblMsg004.Refresh
             flblMsg006.Text                 =    "Sua Versão de "
             flblMsg006.Refresh
             flblMsg007.Text                 =     Format        (fdatVerLoc,    "dd/MM/yyyy")                        & " às " & _
                                                   Format        (fdatVerLoc,    "HH:mm:ss"  )                                 & _
                                                                                 " será substituída pela mais recente de"
             flblMsg007.Refresh
             flblMsg008.Text                 =     Format        (fdatVerRde,    "dd/MM/yyyy")                        & " às " & _
                                                   Format        (fdatVerRde,    "HH:mm:ss"  )                        & "."
             flblMsg008.Refresh
             flblMsg009.Text                 =    "Tecle F3 para Atualizar sua Versão agora."
             flblMsg009.Refresh
        If   fbooArqExe
             flblMsg010.Text                 =    "Após a Atualização, uma outra Execução"
             flblMsg010.Refresh
             flblMsg011.Text                 =    "do Módulo " &  fstrArgmto(1) & " será iniciada já com a nova Versão."
             flblMsg011.Refresh
        End  If
        Else
             flblMsg001.Text                 =    "Programa ativado de forma inadeqada."
             flblMsg001.Refresh
             fcmbF03Atu.Enabled              =     False
        End  If
End     Sub
Private Sub  rlsLimparMensagens
             flblMsg002.Text                 =     ""
             flblMsg002.Refresh
             flblMsg003.Text                 =     ""
             flblMsg003.Refresh
             flblMsg004.Text                 =     ""
             flblMsg004.Refresh
             flblMsg005.Text                 =     ""
             flblMsg005.Refresh
             flblMsg006.Text                 =     ""
             flblMsg006.Refresh
             flblMsg007.Text                 =     ""
             flblMsg007.Refresh
             flblMsg008.Text                 =     ""
             flblMsg008.Refresh
             flblMsg009.Text                 =     ""
             flblMsg009.Refresh
             flblMsg010.Text                 =     ""
             flblMsg010.Refresh
             flblMsg011.Text                 =     ""
             flblMsg011.Refresh
End     Sub
End     Class