Attribute VB_Name = "modTab"
Option Explicit

Private Const mcsModuleName As String = "modTab"
Private Const mcsContainerNotValidError = "the Container object (usually picture box and/or frame) " & "passed to the function listed in 'source' below, must be a control array, " & "and the indexe numbers should correspond to the tab numbers. " & vbCrLf & "(Note: since the tabstrip's index must " & "start at 1, the container control array should start with an index of 1 not 0)"


'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : CreateTabsFromContainerP
' Params : 
'          oTabStrip As TabStrip
'          ocontainer As Object
' Returns: Nothing
' Desc   : The Sub uses parameters oTabStrip As TabStrip and ocontainer As Object for CreateTabsFromContainerP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Sub CreateTabsFromContainerP(oTabStrip As TabStrip, ocontainer As Object)

    
    Const csProcName As String = "HandleTabStripP"
    On Error GoTo PROC_ERR
    
    Dim vitem As Variant
    
    oTabStrip.Tabs.Clear
    'add the tabs according to the 'tag' property of the container control array
    For Each vitem In ocontainer
        oTabStrip.Tabs.Add vitem.Index, , vitem.Tag
    Next
    
    ''if there is only one tab, hide the tab form
    'If oTabStrip.Tabs.Count = 1 Then
    '    oTabStrip.Visible = False
    'End If
    
PROC_EXIT:
    GoSub Proc_Cleanup
    Exit Sub
    
PROC_ERR:
    Dim lErrNum As Long, sErrSrc As String, sErrDesc As String
    sErrSrc = mcsModuleName & "_" & csProcName
    lErrNum = Err.Number
    If lErrNum = 343 Or lErrNum = 35600 Or lErrNum = 438 Then
        'the item is not a control array, so raise the appropriate error
        sErrDesc = mcsContainerNotValidError
    Else
        sErrDesc = Err.Description
    End If
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    
    Err.Raise lErrNum, sErrSrc, sErrDesc
    Exit Sub
    Resume
    
Proc_Cleanup:
    On Error Resume Next
    On Error GoTo 0
    Return

End Sub


'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : HandleTabStripP
' Params : 
'          oTabStrip As TabStrip
'          ocontainer As Object
'          Optional fMoveToFront As Boolean = False
'          Optional vOriginX As Variant
'          Optional vOriginY As Variant
' Returns: Nothing
' Desc   : The Sub uses parameters oTabStrip As TabStrip, ocontainer As Object, Optional fMoveToFront As Boolean = False, Optional vOriginX As Variant and Optional vOriginY As Variant for HandleTabStripP and returns Nothing.
'__________________________________________________
' History
' CDK: 20001112: Added Error Trapping & Comments using
'		Auto-Code Commenter
'__________________________________________________
Public Sub HandleTabStripP(oTabStrip As TabStrip, ocontainer As Object, Optional fMoveToFront As Boolean = False, Optional vOriginX As Variant, Optional vOriginY As Variant)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Comments   : The following function will handle associating a tab
    '               strip's click with the positioning of the appropriate
    '               a frame control array item or picture box control array item
    '             This function assumes that the index of each container item
    '               is the same as it's associated tab
    '               (ie - the tabstrips 1st tab has index of 1.  when this is clicked
    '               the container with the index of 1 will be displayed)
    '             Note: the control array must start at 1 not at 0
    '               since the tab's index must start at 1)
    '             This function COULD be used in conjunction with
    '               the CreateTabsFromContainerP function since it
    '               populates the tabs with the container's index property
    'Parameters : oTabStrip - the tabstrip on that will control
    '               the oContainer
    '             oContainer - the container (usually a pic box or frame)
    '               that will be switched according to the click of oTabStrip
    ' Returns   : Nothing - raises error if parameters are not appropriate
    ' Source    : Charlie Kirkwood
    ' Update    :
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Const csProcName As String = "HandleTabStripP"
    Dim vitem As Variant
    Dim vItemCount As Long
    
    On Error GoTo PROC_ERR
    
    For Each vitem In ocontainer
        vItemCount = vItemCount + 1
    Next
    
    'check to see if the frame is a control array, if we test the index property and it errors to 343
    '   then it's not a control array, the error handler will take care of it
    For Each vitem In ocontainer
    
        If vitem.Index = oTabStrip.SelectedItem.Index Then
            
            'default will be a little past the bottom of the the tabstrip's tabs, but user may
            '   specify origin for pic box
            If IsMissing(vOriginX) Then
                vOriginX = 120
            End If
            
            If IsMissing(vOriginY) Then
                'if the borderstyle of the container is
                '   visible, then account for it,
                '   otherwise push it against the edge
                If vitem.BorderStyle = 0 Then
                    If vItemCount = 1 Then
                        vOriginY = 360
                    Else
                        vOriginY = 420
                    End If
                Else
                    vOriginY = 360
                End If
            End If
            
            vitem.Move oTabStrip.Left + vOriginX, oTabStrip.Top + vOriginY
            
                        
            If fMoveToFront Then
                vitem.ZOrder 0
                vitem.Enabled = True
            End If
            
        Else
        
            'move the container out of view on the form
            vitem.Move -25000, -25000
            vitem.Enabled = False
        
        End If
    
    Next
    
'    lLowerBound = ocontainer.lbound
'    lUpperBound = ocontainer.ubound
'
'    For lCounter = lLowerBound To lUpperBound
'        If lCounter = oTabStrip.SelectedItem.Index Then
'            ocontainer(lCounter).Move oTabStrip.Left + 120, oTabStrip.Top + 480
'        Else
'            'move the container out of view on the form
'            ocontainer(lCounter).Move -25000, -25000
'        End If
'    Next
    
PROC_EXIT:
    GoSub Proc_Cleanup
    Exit Sub
    
PROC_ERR:
    Dim lErrNum As Long, sErrSrc As String, sErrDesc As String
    sErrSrc = mcsModuleName & "_" & csProcName
    If Err.Number = 343 Then
        'the item is not a control array, so raise the appropriate error
        lErrNum = 343
        sErrDesc = mcsContainerNotValidError
    Else
        lErrNum = Err.Number
        sErrDesc = Err.Description
    End If
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise lErrNum, sErrSrc, sErrDesc
    Exit Sub
    
Proc_Cleanup:
    On Error Resume Next
    On Error GoTo 0
    Return

End Sub


