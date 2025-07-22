' =================================================================================
' Author: Gemini
' Date: 2025-07-22
' Description: This module provides a robust framework for programmatically creating
'              a series of geometric shapes (meander lines) and applying a
'              chamfer to a specific edge on each shape. It includes a dedicated
'              subroutine for the picking and chamfering logic to enhance
'              modularity and reusability.
' =================================================================================

' =================================================================================
' Subroutine: ApplyChamferToPickedEdge
' Purpose:    Identifies a specific edge on a given shape by its coordinates and
'             applies a chamfer. This function encapsulates the low-level
'             picking and modification commands.
' Parameters:
'   targetShapeName (String): The full, qualified name of the target shape
'                             (e.g., "component1:mybrick").
'   pickX (Double): The X-coordinate used to identify the target edge/vertex.
'   pickY (Double): The Y-coordinate used to identify the target edge/vertex.
'   pickZ (Double): The Z-coordinate used to identify the target edge/vertex.
'   chamferValue (String): The parameter name or value for the chamfer distance
'                          (e.g., "w_meander_gap").
'   chamferAngle (String): The angle for the chamfer operation (e.g., "45").
' =================================================================================
Sub ApplyChamferToPickedEdge(ByVal targetShapeName As String, _
                             ByVal pickX As Double, _
                             ByVal pickY As Double, _
                             ByVal pickZ As Double, _
                             ByVal chamferValue As String, _
                             ByVal chamferAngle As String)

    ' --- Step 1: Define variables for storing geometric entity IDs ---
    Dim EdgeID As Long
    Dim VertexID As Long
    Dim FaceID As Long

    ' --- Step 2: Programmatically pick entities by coordinates ---
    ' Retrieve the unique IDs of the edge, vertex, and face at the specified location.
    ' Note: The Z-coordinate is adjusted for different picking operations as per the
    ' original logic, which may be required by the CAD software's API.
    EdgeID = Pick.GetEdgeIdFromPoint(targetShapeName, pickX, pickY, pickZ / 2)
    VertexID = Pick.GetVertexIDFromPoint(targetShapeName, pickX, pickY, pickZ)
    FaceID = Pick.GetFaceIdFromPoint(targetShapeName, pickX, pickY / 2, pickZ / 2)

    ' --- Step 3: Validate that the picking operations were successful ---
    If EdgeID > 0 And VertexID > 0 And FaceID > 0 Then
        ' --- Step 4: Construct and log the edge selection command ---
        ' This command emulates a manual user selection of the edge using its ID.
        ' The string is carefully constructed with triple quotes to handle nested quotes.
        Dim pickCmd As String
        pickCmd = "Pick.PickEdgeFromId """ & targetShapeName & """, """ & CStr(EdgeID) & """, """ & CStr(VertexID) & """"
        AddToHistory "Select edge on: " & targetShapeName, pickCmd

        ' --- Step 5: Construct and log the chamfer command (Restored and Corrected) ---
        ' This command applies the chamfer to the previously selected edge.
        ' The command syntax has been restored to its original 4-parameter format.
        ' The FaceID is now correctly included in the command string.
        Dim chamferCmd As String
        chamferCmd = "Solid.ChamferEdge """ & chamferValue & """, """ & chamferAngle & """, ""False"", """ & CStr(FaceID) & """"
        AddToHistory "Chamfer edge on: " & targetShapeName, chamferCmd

        ' --- Step 6: Rebuild the model to apply all pending history list changes ---
        Rebuild
    Else
        ' --- Step 7: Error Handling ---
        ' If any ID is zero or negative, the picking failed. Log an error message.
        Debug.Print "Error: Failed to pick required entities (Edge/Vertex/Face) on shape '" & targetShapeName & "'."
    End If

End Sub

' =================================================================================
' Main Procedure (Main)
' Purpose:    The main entry point of the script. It iteratively creates a series
'             of meander line segments and calls the helper subroutine to
'             chamfer each one, with conditional logic to cap the geometry.
' =================================================================================
Sub Main ()
    ' Define the total number of meander turns to generate.
    Dim turns As Integer
    turns = 3
    ' Loop to create each microstrip unit.
    Dim i As Integer
    For i = 1 To turns
        ' --- Step 1: Define parameters for the current brick segment ---
        Dim xmin As String, xmax As String
        Dim ymin As String, ymax As String
        Dim brickname As String
        Dim history_cmd_brick As String

        ' Construct the coordinate formulas as strings for parametric modeling.
        xmin = "x_patch1-l_patch/2+w_meander*" & CStr(2*i) & "+w_meander_gap*" & CStr(2*i-1)
        xmax = "x_patch1-l_patch/2+w_meander*" & CStr(2*i) & "+w_meander_gap*" & CStr(2*i)
        ymin = "y_patch1-l_patch/2"
        ymax = "y_patch1+l_patch/2-w_chamfer_patch+w_meander*" & CStr(2*i) & "+w_meander_gap*" & CStr(2*i-0.5) & "-w_meander/sqr(2)"
        brickname = "meander_LU2_" & CStr(i)

        ' --- Step 2: Check if the meander segment exceeds the patch boundary ---
        If Evaluate(ymax) <= Evaluate("y_patch1+l_patch/2") Then
            ' --- If within bounds, create the brick and apply the chamfer ---
            history_cmd_brick = "With Brick" & vbCrLf & _
                                "  .Reset" & vbCrLf & _
                                "  .Name """ & brickname & """" & vbCrLf & _
                                "  .Component ""component1""" & vbCrLf & _
                                "  .Material ""PEC""" & vbCrLf & _
                                "  .Xrange """ & xmin & """, """ & xmax & """" & vbCrLf & _
                                "  .Yrange """ & ymin & """, """ & ymax & """" & vbCrLf & _
                                "  .Zrange ""ts"", ""ts+tp""" & vbCrLf & _
                                "  .Create" & vbCrLf & _
                                "End With"
            AddToHistory "Creating brick: " & brickname, history_cmd_brick
            Rebuild

            ' --- Perform the chamfer operation using the dedicated subroutine ---
            Dim shapeName As String
            shapeName = "component1:" & brickname

            Dim pick_x As Double, pick_y As Double, pick_z As Double
            pick_x = Evaluate(xmin)
            pick_y = Evaluate(ymax)
            pick_z = Evaluate("ts+tp/2")

            ApplyChamferToPickedEdge shapeName, pick_x, pick_y, pick_z, "w_meander_gap", "45"
        Else
            ' --- If out of bounds, cap the brick's length and do NOT chamfer ---
            ymax = "y_patch1+l_patch/2-w_meander"

            history_cmd_brick = "With Brick" & vbCrLf & _
                                "  .Reset" & vbCrLf & _
                                "  .Name """ & brickname & """" & vbCrLf & _
                                "  .Component ""component1""" & vbCrLf & _
                                "  .Material ""PEC""" & vbCrLf & _
                                "  .Xrange """ & xmin & """, """ & xmax & """" & vbCrLf & _
                                "  .Yrange """ & ymin & """, """ & ymax & """" & vbCrLf & _
                                "  .Zrange ""ts"", ""ts+tp""" & vbCrLf & _
                                "  .Create" & vbCrLf & _
                                "End With"
            AddToHistory "Creating capped brick: " & brickname, history_cmd_brick
            Rebuild
        End If
    Next i
End Sub
