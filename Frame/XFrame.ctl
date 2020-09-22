VERSION 5.00
Begin VB.UserControl BciFrame 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   HitBehavior     =   2  'Use Paint
   PropertyPages   =   "XFrame.ctx":0000
   ScaleHeight     =   1695
   ScaleWidth      =   3510
   Begin VB.Line Line2 
      X1              =   0
      X2              =   480
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   855
   End
   Begin VB.Label TopLabel 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "BciFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Default Property Values:
Const m_def_Caption = "Frame"
Const m_def_Enabled As Boolean = True
Const m_def_LineVisible As Boolean = True


Const m_def_BackColor = vbWhite
Const m_def_BackHeadColor = vbGrayed
Const m_def_BorderColor = vbBlack
Const m_def_CaptionColor = vbBlack


Const m_def_BorderWidth = 1
Const m_def_BciLineWidth = 1

'Property Variables:
Dim m_Caption As String
Dim m_Enabled As Boolean
Dim m_LineVisible As Boolean

Dim m_BackColor As OLE_COLOR
Dim m_BackHeadColor As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR

Dim m_BorderWidth As Integer
Dim m_BciLineWidth As Integer


'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."











'Properties*****************************************************************


Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = m_Caption
End Property



Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    
:
    lblCaption.Caption = m_Caption
End Property




' enabled code

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
    
    If m_Enabled = True Then
    UserControl.Enabled = True
    Else
    UserControl.Enabled = False
    End If
End Property





' linevisible code

Public Property Get LineVisible() As Boolean
    LineVisible = m_LineVisible
End Property

Public Property Let LineVisible(ByVal vNewValue As Boolean)
    m_LineVisible = vNewValue
    Line2.Visible = m_Enabled
    PropertyChanged "Enabled"
    
    If m_LineVisible = True Then
    Line2.Visible = True
    Else
    Line2.Visible = False
    End If
End Property







 'Back Color code

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property



Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"

    UserControl.BackColor = m_BackColor
End Property






 'Background Head Color code

Public Property Get BackHeadColor() As OLE_COLOR
    BackHeadColor = m_BackHeadColor
End Property



Public Property Let BackHeadColor(ByVal New_BackHeadColor As OLE_COLOR)
    m_BackHeadColor = New_BackHeadColor
    PropertyChanged "BackHeadColor"

    TopLabel.BackColor = m_BackHeadColor
End Property









 'BorderColor code

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property



Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"

   Shape1.BorderColor = m_BorderColor
   Line1.BorderColor = m_BorderColor
   Line2.BorderColor = m_BorderColor
End Property









 'Caption Color code

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = m_CaptionColor
End Property



Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
    m_CaptionColor = New_CaptionColor
    PropertyChanged "CaptionColor"

    lblCaption.ForeColor = m_CaptionColor
End Property









'Border width code

Public Property Get BorderWidth() As Integer
   BorderWidth = m_BorderWidth
End Property



Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    
  
    If m_BorderWidth = 0 Then m_BorderWidth = 1
   Shape1.BorderWidth = m_BorderWidth
End Property






'Line width code

Public Property Get BciLineWidth() As Integer
   BciLineWidth = m_BciLineWidth
End Property



Public Property Let BciLineWidth(ByVal New_BciLineWidth As Integer)
    m_BciLineWidth = New_BciLineWidth
    PropertyChanged "BciLineWidth"
    
  
    If m_BciLineWidth = 0 Then m_BciLineWidth = 1
   Line1.BorderWidth = m_BciLineWidth
End Property






'Font code


Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set lblCaption.Font = vNewValue
    PropertyChanged "Font"
   
End Property











'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
 
    m_Caption = UserControl.Name
    lblCaption.Caption = m_Caption
        
        m_Enabled = m_def_Enabled
        m_LineVisible = m_def_LineVisible


    
    m_BackColor = TopLabel.BackColor
    TopLabel.BackColor = m_BackHeadColor
    
    
     m_BackHeadColor = UserControl.BackColor
    UserControl.BackColor = m_BackColor
     
     
     m_BorderColor = Shape1.BorderColor
    Shape1.BorderColor = m_BorderColor
       Line1.BorderColor = m_BorderColor
       Line2.BorderColor = m_BorderColor

    
       m_CaptionColor = lblCaption.ForeColor
    lblCaption.ForeColor = m_CaptionColor
    
     m_BorderWidth = Shape1.BorderWidth
    Shape1.BorderWidth = m_BorderWidth
    
    
         m_BciLineWidth = Line1.BorderWidth
    Line1.BorderWidth = m_BciLineWidth
    

  UserControl.Enabled = m_Enabled
  Line2.Visible = m_LineVisible
UserControl.BackColor = m_BackColor
TopLabel.BackColor = m_BackHeadColor
lblCaption.Caption = "Frame"
    Set lblCaption.Font = Ambient.Font



    
End Sub







'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
        lblCaption.Caption = m_Caption

    
    
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    
       If m_Enabled = True Then
    UserControl.Enabled = True
    Else
    UserControl.Enabled = False
    End If



   
    m_LineVisible = PropBag.ReadProperty("LineVisible", m_def_LineVisible)
    
      If m_LineVisible = True Then
    Line2.Visible = True
    Else
    Line2.Visible = False
    End If




     m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = m_BackColor

     
     m_BackHeadColor = PropBag.ReadProperty("BackHeadColor", m_def_BackHeadColor)
    TopLabel.BackColor = m_BackHeadColor
    
     m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    Shape1.BorderColor = m_BorderColor
  Line1.BorderColor = m_BorderColor
  Line2.BorderColor = m_BorderColor
      m_CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
    lblCaption.ForeColor = m_CaptionColor
    
    
     m_BorderWidth = PropBag.ReadProperty("BorderWidth", m_def_BorderWidth)
   Shape1.BorderWidth = m_BorderWidth
    
    
      m_BciLineWidth = PropBag.ReadProperty("BciLineWidth", m_def_BciLineWidth)
   Line1.BorderWidth = m_BciLineWidth
    
    
        Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)

    
    
    

   
End Sub











Private Sub UserControl_Resize()
Shape1.Width = ScaleWidth
Shape1.Height = ScaleHeight

lblCaption.Width = ScaleWidth
TopLabel.Width = ScaleWidth

Line1.X2 = ScaleWidth

   Line2.Y1 = UserControl.ScaleHeight - 150
       Line2.Y2 = UserControl.ScaleHeight - 150
        Line2.X2 = ScaleWidth
End Sub










'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("LineVisible", m_LineVisible, m_def_LineVisible)

        Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)

        Call PropBag.WriteProperty("BackHeadColor", m_BackHeadColor, m_def_BackHeadColor)
        
        Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)


        Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)



        Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, m_def_BorderWidth)
        
        
        Call PropBag.WriteProperty("BciLineWidth", m_BciLineWidth, m_def_BciLineWidth)

       Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)


End Sub








