VERSION 5.00
Begin VB.UserControl ShadowedSeperator 
   CanGetFocus     =   0   'False
   ClientHeight    =   15
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ScaleHeight     =   15
   ScaleWidth      =   1005
   ToolboxBitmap   =   "SGLine.ctx":0000
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   1000
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   1000
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ShadowedSeperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************** _
  * SoftGroup Line Control                                              * _
  * SoftGroup Development Corporation - Michael E. Crute                * _
  * mcrute@softgroupcorp.com  -  Chief Software Architect               * _
  * Visit: www.softgroupcorp.com                                        * _
  * Created: 6/6/2002                                                   * _
  * Modified: 6/6/2002                                                  * _
  ***********************************************************************

Private Sub UserControl_Resize()
'This is about as simple as you get. Just extend the line the length of the _
 control and keep the height of the control the same.
Line1.X2 = UserControl.Width
Line2.X2 = UserControl.Width
UserControl.Height = 30
End Sub
