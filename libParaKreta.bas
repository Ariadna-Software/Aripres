Attribute VB_Name = "libParaKreta"
'Para compilar esta verion (SIT NIGEN)
'le quitamos estos datos

'        colkreta
'        kreta2
'        UsuarioHuella
'
'        frmKreta3
'
'
Public GesHuellaDB As BaseDatos2


'    Hay que comentar este trozo
Public ColK2 As ColKreta2    'CON KRETA
'Public ColK2


Public Sub CerrarConexionesKreta()
      
    Set ColK2 = Nothing
    GesHuellaDB.Cerrar
    Set GesHuellaDB = Nothing
End Sub


