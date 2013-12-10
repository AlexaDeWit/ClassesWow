Imports System.IO
'I, Alexa DeWit, 000306848, declare that this is my own original work
'and that I have used no other person's work within. I have also
'not made my work available to anyone else.
Public Class frmScores
    'God Help me, Structs better be passed entirely by value on the stack, as will be primitives
    'Or my entire programming logic needs to be re-thought.
    'I trust myself to understand default pass methods off the top of my head... I hope.



    Private StudentGrades As List(Of StudentData)


    'I am so not dealing with 36 input boxes. I made a class specifically so I don't have to.
    Private StudentInputFields As List(Of StudentFields)
    Private ContainsGoodData As Boolean = False
    Const FILE_NAME As String = "StudentData.csv"
    Private Sub frmScores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StudentGrades = New List(Of StudentData)
        StudentInputFields = New List(Of StudentFields)
        'Add the textboxes to the input field data structure
        'So I never have to type in their names ever ever again
        StudentInputFields.Add(New StudentFields(txtStudentName1, txtStudent1Score1, txtStudent1Score2, txtStudent1Score3, txtStudent1Score4, txtStudent1Score5, txtStudent1Average))
        StudentInputFields.Add(New StudentFields(txtStudentName2, txtStudent2Score1, txtStudent2Score2, txtStudent2Score3, txtStudent2Score4, txtStudent2Score5, txtStudent2Average))
        StudentInputFields.Add(New StudentFields(txtStudentName3, txtStudent3Score1, txtStudent3Score2, txtStudent3Score3, txtStudent3Score4, txtStudent3Score5, txtStudent3Average))
        StudentInputFields.Add(New StudentFields(txtStudentName4, txtStudent4Score1, txtStudent4Score2, txtStudent4Score3, txtStudent4Score4, txtStudent4Score5, txtStudent4Average))
        StudentInputFields.Add(New StudentFields(txtStudentName5, txtStudent5Score1, txtStudent5Score2, txtStudent5Score3, txtStudent5Score4, txtStudent5Score5, txtStudent5Average))
        StudentInputFields.Add(New StudentFields(txtStudentName6, txtStudent6Score1, txtStudent6Score2, txtStudent6Score3, txtStudent6Score4, txtStudent6Score5, txtStudent6Average))

        lblErrorText.Text = vbNullString
    End Sub


    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'reset necessary global data
        lblErrorText.Text = vbNullString
        StudentGrades = New List(Of StudentData)

        'Set up a dataset for one student
        Dim dataSet As StudentData
        'a single grade value used for parsing before adding it to the list
        Dim grade As Double
        Dim gradeTotal As Double = 0.0


        'Clear old data from memory
        StudentGrades.Clear()

        'Start gathering new data
        For Each fieldSet In StudentInputFields

            'clean the dataset from the previous pass
            dataSet = New StudentData
            dataSet.Grades = New List(Of Double)
            gradeTotal = 0.0


            dataSet.Name = fieldSet.StudentName.Text
            For Each textGrade In fieldSet.Grades
                'Attempt to parse a grade
                If (Double.TryParse(textGrade.Text, grade) AndAlso grade >= 0 AndAlso grade <= 100) Then
                    dataSet.Grades.Add(grade)
                    gradeTotal += grade
                Else
                    'Failed a parse, break the sub immediately
                    'Report that the data is not good
                    lblErrorText.Text = "Please enter only numeric grades from 0 to 100."
                    ContainsGoodData = False
                    Exit Sub
                End If
            Next
            dataSet.GradeAverage = gradeTotal / 5.0
            StudentGrades.Add(dataSet)
            fieldSet.ParsedData = dataSet
        Next
        ContainsGoodData = True
        DisplayAverages()
    End Sub
    Private Sub DisplayAverages()
        'Iterate through the text boxes matching by index to display
        For Each fieldSet In StudentInputFields
            fieldSet.GradeAverage.Text = FormatNumber(fieldSet.ParsedData.GradeAverage)
        Next
    End Sub

    Private Sub btnWriteFile_Click(sender As Object, e As EventArgs) Handles btnWriteFile.Click
        'I'm going lazymode and practically copying from the textbook rather than writing my own system.
        If (ContainsGoodData) Then
            Dim studentFile As StreamWriter
            Try
                studentFile = File.CreateText(FILE_NAME)
                'Write each student's data on its own line as comma separated values
                For Each student In StudentGrades
                    studentFile.Write(student.Name + " , ")
                    For Each grade In student.Grades
                        studentFile.Write(CStr(grade) + " , ")
                    Next
                    studentFile.WriteLine(student.GradeAverage)
                Next
                studentFile.Close()
            Catch ex As Exception
                MessageBox.Show("File creation or writing failure.")
            End Try
        Else
            lblErrorText.Text = "No valid data to save. Please calculate first."
        End If

    End Sub

    Private Sub btnReadFile_Click(sender As Object, e As EventArgs) Handles btnReadFile.Click
        'reset necessary global data
        StudentGrades = New List(Of StudentData)
        lblErrorText.Text = vbNullString

        Dim dataSet As StudentData = New StudentData
        Dim studentFile As StreamReader
        'Iterator used for reading in the file
        Dim i As Integer = 0
        'Attempt to read in the student file data
        Try
            studentFile = File.OpenText(FILE_NAME)

            For Each fieldSet In StudentInputFields
                If (TryParseStudentData(studentFile.ReadLine(), dataSet)) Then
                    StudentGrades.Add(dataSet)
                Else
                    MessageBox.Show("Error Reading File. May be corrupted or invalid.")
                    ContainsGoodData = False
                    Exit Sub
                End If
            Next
            studentFile.Close()
        Catch ex As Exception
            MessageBox.Show("File reading failed. Please ensure student data exists before attempting to load it.")
        End Try
        'Must have read 6 items to consider the data good
        If (StudentGrades.Count = 6) Then
            ContainsGoodData = True
            'good data in, set up the associations
            i = 0
            For Each fieldSet In StudentInputFields
                fieldSet.ParsedData = StudentGrades.Item(i)
                fieldSet.StudentName.Text = fieldSet.ParsedData.Name
                For j = 0 To 4
                    fieldSet.Grades.Item(j).Text = FormatNumber(fieldSet.ParsedData.Grades.Item(j))
                Next
                fieldSet.GradeAverage.Text = FormatNumber(fieldSet.ParsedData.GradeAverage)
                i += 1
            Next
        Else
            ContainsGoodData = False
        End If

    End Sub

    Private Function TryParseStudentData(line As String, ByRef target As StudentData) As Boolean
        Dim dataItem As StudentData
        Dim brokenString As String()
        Dim stringSplitPoint As String() = {" , "}
        dataItem.Grades = New List(Of Double)
        'break the CSV into an array
        brokenString = line.Split(stringSplitPoint, StringSplitOptions.None)

        Try
            '
            dataItem.Name = brokenString(0)
            For i = 1 To 5
                dataItem.Grades.Add(CDbl(brokenString(i)))
            Next
            dataItem.GradeAverage = CDbl(brokenString(6))
        Catch ex As Exception
            Return False
        End Try
        target = dataItem
        Return True
    End Function


    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
        ' For this assignement, you are supposed to use the following line
        ' of code to do your printing.
        'PrintDocument1.Print()

        ' PLEASE DO NOT... Instead, please add a Print Preview Dialog to your
        ' form and use the code below instead. This will allow you to use
        ' the on screen print preview and will save the world a bunch of
        ' paper.
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub


    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim OutputString As String = ""
        Dim i As Integer 'iterator

        For Each student In StudentGrades
            OutputString += "[" + student.Name + "]" + vbNewLine
            i = 0
            For Each grade In student.Grades
                i += 1
                OutputString += "Test " + CStr(i) + ": " + FormatNumber(grade) + vbNewLine
            Next
            OutputString += "Average: " + FormatNumber(student.GradeAverage) + vbNewLine
            'some padding between students
            OutputString += vbNewLine + vbNewLine
        Next

        e.Graphics.DrawString(OutputString, New Font("Courier New", 14), Brushes.Blue, 100, 100)
    End Sub
End Class
'This Class exists to allow me to deal with the data by groups instead of a massive unrepeating mess
Public Class StudentFields
    Public StudentName As TextBox
    Public Grades As List(Of TextBox)
    Public GradeAverage As TextBox
    'Allow the fields to reference their parsed and calculated data
    'this is to ensure that a fieldset correctly corresponds to its data at all times
    Public ParsedData As StudentData
    Public Sub New(ByRef student As TextBox, ByRef grade1 As TextBox, ByRef grade2 As TextBox, ByRef grade3 As TextBox, ByRef grade4 As TextBox, ByRef grade5 As TextBox, ByRef average As TextBox)
        Grades = New List(Of TextBox)
        StudentName = student
        Grades.Add(grade1)
        Grades.Add(grade2)
        Grades.Add(grade3)
        Grades.Add(grade4)
        Grades.Add(grade5)
        GradeAverage = average
    End Sub
End Class
'Obligatory Struct I was required to make
Public Structure StudentData
    Public Name As String
    Public Grades As List(Of Double)
    Public GradeAverage As Double
End Structure
