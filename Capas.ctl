VERSION 5.00
Begin VB.UserControl Capas 
   BackStyle       =   0  'Transparent
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   8535
   Begin VB.Image Image1 
      Height          =   8520
      Left            =   90
      Stretch         =   -1  'True
      Top             =   90
      Width           =   8310
   End
End
Attribute VB_Name = "Capas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Habilitado = 0
'Property Variables:
Dim m_Habilitado As Boolean
'Event Declarations:
Event Clique() 'MappingInfo=Image1,Image1,-1,DblClick
Attribute Clique.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Foto() As Picture
Attribute Foto.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Foto = Image1.Picture
End Property

Public Property Set Foto(ByVal New_Foto As Picture)
    Set Image1.Picture = New_Foto
    PropertyChanged "Foto"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Habilitado() As Boolean
    Habilitado = m_Habilitado
End Property

Public Property Let Habilitado(ByVal New_Habilitado As Boolean)
    m_Habilitado = New_Habilitado
    PropertyChanged "Habilitado"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Habilitado = m_def_Habilitado
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Foto", Nothing)
    m_Habilitado = PropBag.ReadProperty("Habilitado", m_def_Habilitado)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Foto", Picture, Nothing)
    Call PropBag.WriteProperty("Habilitado", m_Habilitado, m_def_Habilitado)
End Sub

Private Sub Image1_DblClick()
    RaiseEvent Clique
End Sub

