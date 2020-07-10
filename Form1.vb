Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.OleDb

Imports System.IO
Public Class Form1
    Dim baris As Integer
    Dim iterasi As Integer
    ' Dim baris As Integer
    Dim tipeA As Integer = 3
    Dim l As ListViewItem
    Dim P_Stack As String
    Dim N As Integer
    Dim x As String
    Dim Tc As String
    Dim P_H2O As String
    Dim pp_H2 As String
    Dim pp_O2 As String
    ' Dim e As Exception
    Dim q As Integer
    Dim V_out As String
    Dim z As Integer
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        TabControl1.SelectedIndex = 2
        wait(1)
        Button29.PerformClick()
        wait(1)
        TabControl1.SelectedIndex = 3
        Button25.PerformClick()

    End Sub
    Public Sub ControlBmpToFile(ByVal control As Control, ByVal file As String)
        Dim bmp As New Bitmap(control.Width, control.Height)
        control.DrawToBitmap(bmp, control.DisplayRectangle)

        bmp.Save(file, System.Drawing.Imaging.ImageFormat.Jpeg)
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click



        For Me.baris = 1 To 26
            l = Me.ListView1.Items.Add("")
            For j As Integer = 1 To Me.ListView1.Columns.Count
                l.SubItems.Add("")
            Next
        Next
        Me.ListView1.Items(0).SubItems(1).Text = TextBox1.Text
        Me.ListView1.Items(0).SubItems(0).Text = Label1.Text
        Me.ListView1.Items(0).SubItems(2).Text = Label3.Text

        Me.ListView1.Items(1).SubItems(1).Text = TextBox2.Text
        Me.ListView1.Items(1).SubItems(0).Text = Label2.Text
        Me.ListView1.Items(1).SubItems(2).Text = Label4.Text

        Me.ListView1.Items(2).SubItems(1).Text = TextBox3.Text
        Me.ListView1.Items(2).SubItems(0).Text = Label5.Text
        Me.ListView1.Items(2).SubItems(2).Text = Label6.Text

        Me.ListView1.Items(3).SubItems(1).Text = TextBox4.Text
        Me.ListView1.Items(3).SubItems(0).Text = Label7.Text
        Me.ListView1.Items(3).SubItems(2).Text = Label8.Text

        Me.ListView1.Items(4).SubItems(1).Text = TextBox5.Text
        Me.ListView1.Items(4).SubItems(0).Text = Label9.Text
        Me.ListView1.Items(4).SubItems(2).Text = Label10.Text

        Me.ListView1.Items(5).SubItems(1).Text = TextBox6.Text
        Me.ListView1.Items(5).SubItems(0).Text = Label11.Text
        Me.ListView1.Items(5).SubItems(2).Text = Label12.Text

        Me.ListView1.Items(6).SubItems(1).Text = TextBox7.Text
        Me.ListView1.Items(6).SubItems(0).Text = Label13.Text
        Me.ListView1.Items(6).SubItems(2).Text = Label16.Text

        Me.ListView1.Items(7).SubItems(1).Text = TextBox8.Text
        Me.ListView1.Items(7).SubItems(0).Text = Label14.Text
        Me.ListView1.Items(7).SubItems(2).Text = Label15.Text

        Me.ListView1.Items(8).SubItems(1).Text = TextBox9.Text
        Me.ListView1.Items(8).SubItems(0).Text = Label17.Text
        Me.ListView1.Items(8).SubItems(2).Text = ""

        Me.ListView1.Items(9).SubItems(1).Text = TextBox10.Text
        Me.ListView1.Items(9).SubItems(0).Text = Label19.Text
        Me.ListView1.Items(9).SubItems(2).Text = ""

        Me.ListView1.Items(10).SubItems(1).Text = TextBox23.Text
        Me.ListView1.Items(10).SubItems(0).Text = Label41.Text
        Me.ListView1.Items(10).SubItems(2).Text = ""

        Me.ListView1.Items(11).SubItems(1).Text = TextBox11.Text ^ TextBox12.Text
        Me.ListView1.Items(11).SubItems(0).Text = Label20.Text
        Me.ListView1.Items(11).SubItems(2).Text = Label21.Text

        Me.ListView1.Items(12).SubItems(1).Text = TextBox13.Text
        Me.ListView1.Items(12).SubItems(0).Text = Label23.Text
        Me.ListView1.Items(12).SubItems(2).Text = Label24.Text

        Me.ListView1.Items(13).SubItems(1).Text = TextBox14.Text
        Me.ListView1.Items(13).SubItems(0).Text = Label25.Text
        Me.ListView1.Items(13).SubItems(2).Text = Label26.Text

        Me.ListView1.Items(14).SubItems(1).Text = TextBox15.Text
        Me.ListView1.Items(14).SubItems(0).Text = Label27.Text
        Me.ListView1.Items(14).SubItems(2).Text = Label28.Text

        Me.ListView1.Items(15).SubItems(1).Text = TextBox16.Text
        Me.ListView1.Items(15).SubItems(0).Text = Label29.Text
        Me.ListView1.Items(15).SubItems(2).Text = ""

        Me.ListView1.Items(16).SubItems(1).Text = TextBox17.Text
        Me.ListView1.Items(16).SubItems(0).Text = Label30.Text
        Me.ListView1.Items(16).SubItems(2).Text = ""

        Me.ListView1.Items(17).SubItems(1).Text = TextBox18.Text
        Me.ListView1.Items(17).SubItems(0).Text = Label31.Text
        Me.ListView1.Items(17).SubItems(2).Text = Label32.Text

        Me.ListView1.Items(18).SubItems(1).Text = TextBox19.Text
        Me.ListView1.Items(18).SubItems(0).Text = Label33.Text '
        Me.ListView1.Items(18).SubItems(2).Text = Label34.Text

        Me.ListView1.Items(19).SubItems(1).Text = TextBox20.Text
        Me.ListView1.Items(19).SubItems(0).Text = Label35.Text
        Me.ListView1.Items(19).SubItems(2).Text = Label36.Text

        Me.ListView1.Items(20).SubItems(1).Text = TextBox21.Text
        Me.ListView1.Items(20).SubItems(0).Text = Label37.Text
        Me.ListView1.Items(20).SubItems(2).Text = Label38.Text

        Me.ListView1.Items(21).SubItems(1).Text = TextBox22.Text
        Me.ListView1.Items(21).SubItems(0).Text = Label39.Text
        Me.ListView1.Items(21).SubItems(2).Text = Label40.Text

        Me.ListView1.Items(22).SubItems(1).Text = TextBox26.Text
        Me.ListView1.Items(22).SubItems(0).Text = Label47.Text
        Me.ListView1.Items(22).SubItems(2).Text = Label46.Text
    End Sub
    Public Sub wait(ByVal Dt As Double)
        Dim IDay As Double = Date.Now.DayOfYear
        Dim CDay As Double
        Dim ITime As Double = Date.Now.TimeOfDay.TotalSeconds
        Dim CTime As Double
        Dim DiffDay As Double
        Try
            Do
                Application.DoEvents()
                CDay = Date.Now.DayOfYear
                CTime = Date.Now.TimeOfDay.TotalSeconds
                DiffDay = CDay - IDay
                CTime = CTime + 86400 * DiffDay
                If CTime >= ITime + Dt Then Exit Do

            Loop
        Catch e As Exception
        End Try
    End Sub
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text
        
        Dim i As String
        Dim P_air As String = TextBox5.Text
        Dim b As String
        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text
        Dim V_act As String
        Dim io As String = TextBox11.Text ^ TextBox12.Text
        Dim V_ohmic As String
        Dim rin As String = TextBox8.Text
        Dim term As String
        Dim Bt As String = TextBox13.Text
        Dim V_conc As String
        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text
        Dim E_nernst As String
        Dim a As Integer
        Chart1.Series(0).Points.Clear()
        ListView2.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 
        
        N = TextBox24.Text / TextBox26.Text
        z = TextBox25.Text

        For Me.baris = z + 1 To N 'Step 0.01

            l = Me.ListView2.Items.Add("")
            For j As Integer = 1 To Me.ListView2.Columns.Count
                l.SubItems.Add("")
            Next
            Try
                For Me.iterasi = 2 To tipeA
                    ListView2.Items(baris - 1).SubItems(1).Text = CDec((baris - 1) * TextBox26.Text)
                    ListView2.Items(baris - 1).SubItems(0).Text = baris

                    i = CStr((baris - 1) * TextBox26.Text) '* Math.Sqrt(2)))
                    'Calculation of Partial Pressures 
                    'Calculation of saturation pressure of water 

                    x = -2.1794 + (0.02953 * CDec(Tc)) - CDec(9.1837 * (10 ^ -5) * (Tc ^ 2)) + (1.4454 * (10 ^ -7) * (CDec(Tc) ^ 3))
                    P_H2O = (10 ^ x)
                    'Calculation of partial pressure of hydrogen 
                    pp_H2 = (0.5 * CDec((P_H2) / (Math.Exp(1.653 * i / (tk ^ 1.334))) - P_H2O))
                    'Calculation of partial pressure of oxygen 
                    pp_O2 = (P_air / Math.Exp(4.192 * i / (tk ^ 1.334))) - P_H2O
                    'Activation Losses 
                    b = (R * tk / (2 * alpha * F))
                    V_act = (-b * Math.Log(i / io))
                    'Tafel equation 

                    'Ohmic Losses
                    V_ohmic = -(i * rin)

                    'Mass Transport Losses 
                    term = (1.5 - (Bt * i))
                    If term > 0 Then
                        V_conc = Alpha1 * (i ^ k) * Math.Log(term)
                    Else
                        V_conc = 0
                    End If


                    'Calculation of Nernst voltage 
                    E_nernst = CDec(-Gf_liq / (2 * F)) + (((R * tk) / (2 * F)) * Math.Log(P_H2O / (pp_H2 * (pp_O2 ^ 0.5))))
                    'Calculation of output voltage 


                    If i = 0 Then
                        '
                        E_nernst = CDec((-Gf_liq / (2 * F)) - ((R * tk) * (1 + Math.Log(pp_H2 * (pp_O2 ^ 0.5))) / (2 * F)))
                        'i = 0.001
                        V_out = CDec(E_nernst) - i * rin + Val(V_act) + Val(V_conc)
                        ListView2.Items(baris - 1).SubItems(2).Text = V_out
                    Else
                        V_out = CDec(E_nernst) + CDec(V_ohmic) + CDec(V_act) + CDec(V_conc)
                        ListView2.Items(baris - 1).SubItems(2).Text = CDec((V_out)) '+ Math.Abs(Val(ListView2.Items(baris - 1).SubItems(2).Text) - Val(ListView2.Items(baris - 2).SubItems(2).Text)) / Math.Sqrt(2)
                    End If




                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        N = baris
                        'Exit For
                        ListView2.Items(baris - 1).Remove()
                        ' ListView2.Items.Clear()

                        ' Exit Sub
                    Else
                        If baris > 1 Then
                            ListView2.Items(baris - 1).SubItems(3).Text = CDec(ListView2.Items(baris - 2).SubItems(2).Text - ListView2.Items(baris - 1).SubItems(2).Text)
                        Else
                            ListView2.Items(baris - 1).SubItems(3).Text = CDec(0)

                        End If
                        ' If Val(ListView2.Items(baris - 1).SubItems(2).Text) < 0 Then
                        'V_out = Val(0)
                        ' ListView2.Items(baris - 1).SubItems(2).Text = V_out
                        'ListView2.Items(baris - 2).SubItems(2).Text = V_out
                        'N = baris - 1
                        ' End If



                        If ComboBox12.Text = "Real Time" Then

                            Chart1.Series(0).Points.AddXY(CDec(ListView2.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView2.Items(baris - 1).SubItems(2).Text))
                            If ComboBox17.Text = "Point" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                            ElseIf ComboBox17.Text = "Bar" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                            ElseIf ComboBox17.Text = "Area" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                            ElseIf ComboBox17.Text = "Fast Line" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                            Else
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                            End If

                            If ComboBox13.Text = "Red" Then
                                Chart1.Series(0).Color = Color.Red
                            ElseIf ComboBox13.Text = "Green" Then
                                Chart1.Series(0).Color = Color.Green
                            ElseIf ComboBox13.Text = "Blue" Then
                                Chart1.Series(0).Color = Color.Blue
                            Else
                                Chart1.Series(0).Color = Color.Brown
                            End If
                            If ComboBox14.Text = "Dash" Then
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            Else
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            End If

                            wait(0.001)
                        End If
                        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
                        Chart1.ChartAreas("ChartArea1").AxisY.Title = "Output Voltage (V)"
                        'If ListView2.Items(baris - 1).SubItems(3).Text < 0 Then
                        'End

                        'End If

                    End If


                Next
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Catch t As Exception
            End Try
        Next
        For a = 1 To ListView2.Items.Count
            If ListView2.Items(CInt(ListView2.Items.Count) - 1).SubItems(0).Text = "" Then
                ListView2.Items(CInt(ListView2.Items.Count) - 1).Remove()
            End If
            'Next
            'ListView2.Items(baris - 1).Remove()

            'Exit Sub
        Next
       
        ' Button14.PerformClick()
       

        'MsgBox(Chart1.ChartAreas(0).AxisX.ScaleView.Size)

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ListView2.Items.Clear()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim i As Integer


        TextBox35.Text = ListView2.Items(0).SubItems(2).Text
        Dim g As Integer
        'Dim 1 as index
        For g = 2 To ListView2.Items.Count

            If (ListView2.Items(g - 2).SubItems(2).Text) < (ListView2.Items(g - 1).SubItems(2).Text) And TextBox35.Text < (ListView2.Items(g - 1).SubItems(2).Text) Then
                TextBox35.Text = (ListView2.Items(g - 1).SubItems(2).Text)
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)

            Else
                TextBox35.Text = TextBox35.Text
                ''  End If

                ' End If

            End If

        Next

        TextBox34.Text = ListView2.Items(0).SubItems(2).Text
        'Dim i As Integer
        For i = 2 To ListView2.Items.Count

            'jika angka pertama lebih kecil dari pada angka kedua dan textbox 34 lebih besar dari pada angka kedua maka
            'nilai terkecil adalah di textbox34

            If CDec(ListView2.Items(i - 1).SubItems(2).Text) < CDec(ListView2.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(ListView2.Items(i - 1).SubItems(2).Text) Then
                TextBox34.Text = CDec(ListView2.Items(i - 1).SubItems(2).Text)
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)

                ' ListView2.Items(i - 2).BackColor = Color.Red
            Else
                TextBox34.Text = Val(TextBox34.Text)
                ''  End If

                ' End If

            End If

        Next i
        TextBox33.Text = TextBox35.Text - TextBox34.Text
        TextBox45.Text = ListView2.Items(ListView2.Items.Count - 1).SubItems(1).Text


        TextBox31.Text = Math.Abs(CDec(ListView2.Items(0).SubItems(3).Text))

        'Dim 1 as index
        For g = 2 To ListView2.Items.Count

            If Math.Abs(CDec(ListView2.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(ListView2.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(ListView2.Items(g - 1).SubItems(3).Text)) Then
                TextBox31.Text = Math.Abs(CDec(ListView2.Items(g - 1).SubItems(3).Text))
                ' TextBox44.BackColor = Color.LightBlue
                'TextBox44.Text = Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000")
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)
                ' ListView2.Items(g - 1).BackColor = Color.Blue
                ' ListView2.Items(g - 2).BackColor = Color.Blue
                If ListView2.Items(g - 1).SubItems(3).Text < 0 Then
                    TextBox31.Text = Val("-" + TextBox31.Text)
                End If
            Else
                If ListView2.Items(g - 1).SubItems(3).Text < 0 Then
                    TextBox31.Text = Val("-" + TextBox31.Text)

                Else
                    TextBox31.Text = Val(TextBox31.Text)
                End If
                ''  End If

                ' End If

            End If

        Next g


        TextBox29.Text = Math.Abs(CDec(ListView2.Items(1).SubItems(3).Text))
        'Dim i As Integer

        For i = 3 To ListView2.Items.Count

            'jika angka pertama lebih kecil dari pada angka kedua dan textbox 7 lebih besar dari pada angka kedua maka
            'nilai terkecil adalah di textbox7
            If Math.Abs(CDec(ListView2.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(ListView2.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(ListView2.Items(i - 1).SubItems(3).Text)) Then
                TextBox29.Text = Math.Abs(CDec(ListView2.Items(i - 2).SubItems(3).Text))
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)
                ' ListView2.Items(i - 1).BackColor = Color.Yellow
                ' ListView2.Items(i - 2).BackColor = Color.Yellow
                ' If ListView2.Items(i - 1).SubItems(3).Text < 0 Then
                'TextBox29.Text = Val("-" + TextBox29.Text)
                TextBox43.BackColor = Color.Yellow
                TextBox43.Text = Format(Val(ListView2.Items(i - 2).SubItems(2).Text), "0.000")
                ' End If
            Else
                ' If ListView2.Items(i - 1).SubItems(3).Text < 0 Then
                ' TextBox29.Text = Val("-" + TextBox29.Text)
                '   Else
                TextBox29.Text = Val(TextBox29.Text)
                ' End If

                ''  End If

                ' End If

            End If

        Next i

        'Dim maximum As String
        'Dim minimum As String
        ' maximum = (TextBox35.Text + (TextBox28.Text))
        '  minimum = (TextBox31.Text - (TextBox28.Text))
        ' Dim mini As String

        ' If baris > 2 Then

        'End If
        '  Chart1.ChartAreas(0).AxisY.Minimum = TextBox35.Text 'Then
        '  Chart1.ChartAreas(0).AxisY.Maximum = TextBox34.Text
        'Chart1.AutoScaling = True
        'Chart1.Series.Dispose()

        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
        Chart1.ChartAreas("ChartArea1").AxisY.Title = "Output Voltage (V)"
        ' Chart1.ChartAreas(0).AxisX.ScaleView.Size = Val(baris)
        ' Chart1.ChartAreas(0).AxisY.ScaleView.Size = Val(TextBox33.Text)
        wait(0.001)
        TextBox30.Text = "0"
        For i = 1 To ListView2.Items.Count
            TextBox30.Text = TextBox30.Text + CDec(ListView2.Items(i - 1).SubItems(2).Text)
        Next i
        TextBox30.Text = TextBox30.Text / ListView2.Items.Count
        For g = 1 To ListView2.Items.Count
            If CDec(ListView2.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                ListView2.Items(g - 1).BackColor = Color.LightGreen
                TextBox35.BackColor = Color.LightGreen


            ElseIf ListView2.Items(g - 1).SubItems(3).Text = Val(TextBox29.Text) Then
                ListView2.Items(g - 1).BackColor = Color.Yellow
                TextBox29.BackColor = Color.Yellow



            ElseIf Val(ListView2.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                ListView2.Items(g - 1).BackColor = Color.LightBlue
                TextBox31.BackColor = Color.LightBlue

            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox34.Text Then


                ListView2.Items(g - 1).BackColor = Color.Orange
                TextBox34.BackColor = Color.Orange

            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                ListView2.Items(g - 1).BackColor = Color.LightCyan
                TextBox30.BackColor = Color.LightCyan

            End If

        Next g

        'Button19.PerformClick()
        ' Chart1.ChartAreas(0).AxisY.Interval = TextBox33.Text / 4
        ' Chart1.ChartAreas(0).AxisX.Interval = TextBox24.Text / 5

        Chart1.ChartAreas(0).AxisX.LineWidth = 1
        'Chart1.ChartAreas(0).BorderDashStyle.DashDot = True
        'Chart1.ChartAreas(0).BackImageAlignment.linewidth = 0
        ' Chart1.ChartAreas(0).AxisY.LabelStyle.Enabled = False
        'Button32.PerformClick()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = TextBox24.Text / 2
        Chart1.ChartAreas(0).AxisY.ScaleView.Size = TextBox33.Text / 2
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = TextBox24.Text * 2
        Chart1.ChartAreas(0).AxisY.ScaleView.Size = TextBox33.Text * 2
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet

        End With
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub
    
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Chart1.Series(0).Points.Clear()
        ' Button14.PerformClick()
        ' wait(1)
        Dim baris As Integer
        
        For baris = 1 To ListView2.Items.Count
            ' If ComboBox12.Text = "Real Time" Then
            Chart1.Series(0).Points.AddXY(CDec(ListView2.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView2.Items(baris - 1).SubItems(2).Text))
            ' Chart1.Series(0).Points.AddXY(CDec(ListView2.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView2.Items(baris - 1).SubItems(2).Text))
            If ComboBox17.Text = "Point" Then
                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
            ElseIf ComboBox17.Text = "Bar" Then
                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
            ElseIf ComboBox17.Text = "Area" Then
                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
            ElseIf ComboBox17.Text = "Fast Line" Then
                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

            Else
                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
            End If

            ' Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline


            If ComboBox13.Text = "Red" Then
                Chart1.Series(0).Color = Color.Red
            ElseIf ComboBox13.Text = "Green" Then
                Chart1.Series(0).Color = Color.Green
            ElseIf ComboBox13.Text = "Blue" Then
                Chart1.Series(0).Color = Color.Blue
            Else
                Chart1.Series(0).Color = Color.Brown
            End If
            If ComboBox14.Text = "Dash" Then
                With Chart1.ChartAreas(0)
                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                End With
            Else
                With Chart1.ChartAreas(0)
                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                End With
            End If

            'wait(0.001)
            'End If

        Next


        Chart1.ChartAreas(0).AxisX.Title = "Current density (A/cm^2)"
        Chart1.ChartAreas(0).AxisY.Title = "Output Voltage (V)"
        ' Chart1.ChartAreas(0).RecalculateAxesScale()
        ' Button14.PerformClick()
        ' Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        'Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        'Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        'Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        'Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        'Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        'Chart1.ChartAreas(0).AxisX.IsMarginVisible = True
        'Chart1.ChartAreas(0).RecalculateAxesScale()
        ' Chart1.BorderlineWidth = 0
        'Me.Chart1.ChartAreas(0).AxisY.IsStartedFromZero = False
        ' Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(Val(TextBox35.Text), "0.0") + TextBox33.Text / 10 + TextBox28.Text
        '  If TextBox34.Text < 0 Then
        'Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(Val(TextBox34.Text), "0.0") + TextBox34.Text / 10 - TextBox28.Text
        ' Button14.PerformClick()
        ' Else
        ' Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(Val(TextBox34.Text), "0.0") - TextBox34.Text / 10 - TextBox28.Text
        ' End If
        Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
        'Me.Chart1.ChartAreas("ChartArea1").AxisY.IsReversed = False
        ' Chart1.ChartAreas(0).AxisX.IntervalOffset = 1
        'Chart1.ChartAreas(0).AxisY.IntervalOffset = 1
        Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
        
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (TextBox45.Text - 0) / 5), "0.00E0")
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
        'Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = TextBox24.Text / 100 + TextBox27.Text
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (TextBox45.Text - 0) / 100), "0.00E0")
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
        'Me.Chart1.ChartAreas(0).AxisX.IsReversed = False
        Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
        Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
        'Chart1.ChartAreas(0).AxisX.IntervalOffset = TextBox24.Text / 1000 + TextBox27.Text
        'Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "MM"
        'Chart1.ChartAreas(0).AxisY.Interval = TextBox33.Text / 4
        ' Chart1.ChartAreas(0).AxisX.Interval = (TextBox25.Text + 2 * TextBox24.Text / 1000 + 2 * TextBox27.Text) / 5
        'Chart1.Series("Series1").AxisLabel = True
        ' MsgBox("finish")

    End Sub


    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Dim saveFileDialog1 As New SaveFileDialog()

        ' Sets the current file name filter string, which determines 
        ' the choices that appear in the "Save as file type" or 
        ' "Files of type" box in the dialog box.
        saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        ' Set image file format
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As ChartImageFormat = ChartImageFormat.Bmp

            If saveFileDialog1.FileName.EndsWith("bmp") Then
                format = ChartImageFormat.Bmp
            Else
                If saveFileDialog1.FileName.EndsWith("jpg") Then
                    format = ChartImageFormat.Jpeg
                Else
                    If saveFileDialog1.FileName.EndsWith("emf") Then
                        format = ChartImageFormat.Emf
                    Else
                        If saveFileDialog1.FileName.EndsWith("gif") Then
                            format = ChartImageFormat.Gif
                        Else
                            If saveFileDialog1.FileName.EndsWith("png") Then
                                format = ChartImageFormat.Png
                            Else
                                If saveFileDialog1.FileName.EndsWith("tif") Then
                                    format = ChartImageFormat.Tiff
                               
                                End If
                            End If ' Save image
                        End If
                    End If
                End If
            End If
            Chart1.SaveImage(saveFileDialog1.FileName, format)
        End If
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        For g = 1 To ListView2.Items.Count
            If ListView2.Items(g - 1).BackColor = Color.LightGreen Then
                'ListView2.Items(g - 1).BackColor = Color.LightGreen
                'TextBox35.Text = Format(Val(TextBox35.Text), "0.000")
                Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000") & " Volts"

                'ChartArea1.AxisX.islabelautofit = True
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                '  
            ElseIf ListView2.Items(g - 1).BackColor = Color.Yellow Then
                ' ListView2.Items(g - 1).BackColor = Color.Yellow

                
                Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000") & " Volts"
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow


            ElseIf Val(ListView2.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                ' ListView2.Items(g - 1).BackColor = Color.LightBlue
                ' TextBox31.BackColor = Color.LightBlue
                
                Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000") & " Volts"
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                ' Chart1.Series(0).YValueMembers = "Quan"
                'Chart1.Series(0).Points(g).LabelForeColor = Color.LightBlue
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                'MsgBox(g)
                'MsgBox(ListView2.Items(g - 1).SubItems(2).Text)

                'ListView2.Items(g - 1).BackColor = Color.Orange
                ' TextBox34.BackColor = Color.Orange
                Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000") & " Volts"
                ' MsgBox(Chart1.Series(0).Points(g - 1).Label)
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                'Chart1.Series(0).Points(g).LabelForeColor = Color.Orange
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                ' ListView2.Items(g - 1).BackColor = Color.LightCyan
                ' TextBox30.BackColor = Color.LightCyan
                Chart1.Series(0).Points(g - 1).Label = "mean = " + ListView2.Items(g - 1).SubItems(2).Text & " Volts"
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                'Chart1.Series(0).Points(g).LabelForeColor = Color.LightCyan

            End If

        Next g

        Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
        Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
        Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
        'Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopRight
        ' Button15.PerformClick()
    End Sub

    Private Sub Chart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView2.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView2.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView2.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView2.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView2.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView2.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView2.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub
    ' End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        Chart1.Series(0).Points.Clear()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With

    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        With Chart1.ChartAreas(0)
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        For g = 1 To ListView2.Items.Count
            If CDec(ListView2.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                'ListView2.Items(g - 1).BackColor = Color.LightGreen
                TextBox35.Text = Format(Val(TextBox35.Text), "0.000")
                Chart1.Series(0).Points(g - 1).Label = ""
                'Chart1.Series(0).Points(g - 1).LabelFormat.ToLower()

                Chart1.Series(0).Points(g - 1).MarkerSize = 0
               
            ElseIf ListView2.Items(g - 1).BackColor = Color.Yellow Then

                TextBox43.BackColor = Color.Yellow
                TextBox43.Text = Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000")
                Chart1.Series(0).Points(g - 1).Label = ""
                Chart1.Series(0).Points(g - 1).MarkerSize = 0
               


            ElseIf Val(ListView2.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                ' ListView2.Items(g - 1).BackColor = Color.LightBlue
                ' TextBox31.BackColor = Color.LightBlue
                '  TextBox44.BackColor = Color.LightBlue
                'TextBox44.Text = Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.000")
                Chart1.Series(0).Points(g - 1).Label = ""

                Chart1.Series(0).Points(g - 1).MarkerSize = 0
                
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                
                Chart1.Series(0).Points(g - 1).Label = ""
                
                Chart1.Series(0).Points(g - 1).MarkerSize = 0
                
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                
                Chart1.Series(0).Points(g - 1).Label = ""

                Chart1.Series(0).Points(g - 1).MarkerSize = 0
            End If

        Next g
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim openFileDialog1 As New OpenFileDialog()
        'set the root to the z drive
        openFileDialog1.InitialDirectory = "D:\"
        'make sure the root goes back to where the user started
        ''   openFileDialog1.RestoreDirectory = True
        'show the dialog
        ''  openFileDialog1.ShowDialog()
        ''
        '' If (openFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then
        ''Path = openFileDialog1.FileName
        ''  End If

        ' Call ShowDialog.
        Dim result As DialogResult = openFileDialog1.ShowDialog()

        ' Test result.
        Dim path As String = openFileDialog1.FileName
        If result = Windows.Forms.DialogResult.OK Then
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=Excel 12.0;"
            ' Get the file name.

            '  Try
            Dim table As DataTable = New DataTable()
            Dim excelName As String = "Sheet1"
            Dim strConnection As String = String.Format(connStr)
            Dim conn As OleDbConnection = New OleDbConnection(strConnection)
            conn.Open()
            Dim oada As OleDbDataAdapter = New OleDbDataAdapter("select * from [" & excelName & "$]", strConnection)
            table.TableName = "TableInfo"
            oada.Fill(table)
            conn.Close()
            ListView2.Items.Clear()

            For i As Integer = 0 To table.Rows.Count - 1
                Dim drow As DataRow = table.Rows(i)

                If drow.RowState <> DataRowState.Deleted Then
                    Dim lvi As ListViewItem = New ListViewItem(drow("No").ToString())
                    lvi.SubItems.Add(drow("Curr density (A/cm^2)").ToString())
                    lvi.SubItems.Add(drow("Output voltage (Volts)").ToString())
                    lvi.SubItems.Add(drow("Sequence Discreapancy").ToString())
                    ListView2.Items.Add(lvi)
                End If
            Next
            ' Read in text.
            '' Dim text As String = File.ReadAllText(path)

            ' For debugging.
            ''  Me.Text = text.Length.ToString

            '  Catch ex As Exception

            ' Report an error.
            ''   Me.Text = "Error"

            ' End Try
        End If


        ' For Me.baris = 1 To Me.ListView2.Items.Count
        ' l = Me.ListView2.Items.Add("")
        ' For j As Integer = 1 To Me.ListView2.Columns.Count
        ' l.SubItems.Add("")
        'Next
        'ListView2.Items(baris - 1).SubItems(0).Text = baris
        ' Next


    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        With Chart1.ChartAreas(0)
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox29.Text = ""
        TextBox29.BackColor = Color.White
        TextBox30.Text = ""
        TextBox30.BackColor = Color.White
        TextBox31.Text = ""
        TextBox31.BackColor = Color.White
        TextBox33.Text = ""
        TextBox33.BackColor = Color.White
        TextBox34.Text = ""
        TextBox34.BackColor = Color.White
        TextBox35.Text = ""
        TextBox35.BackColor = Color.White
        'TextBox44.BackColor = Color.White
        ' TextBox44.Text = ""
        TextBox43.Text = ""
        TextBox43.BackColor = Color.White

    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Button39.PerformClick()
        Button6.PerformClick()
        Button34.PerformClick()
        MsgBox("All data have cleared")
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        ListView3.Items.Clear()
        ' Dim i As Integer
        If Me.TextBox35.Text = "" Then
            MsgBox("Run FuelCell Polarization First")
        Else
            q = 0
            For i = 1 To ListView2.Items.Count
                If Not ListView2.Items(i - 1).SubItems(0).Text = "" Then
                    q = q + 1
                Else
                    q = q
                End If
            Next
            wait(0.001)
            'N = ListView2.Items.Count
            z = TextBox25.Text
            For Me.baris = z + 1 To q 'Step 0.01
                l = Me.ListView3.Items.Add("")
                For j As Integer = 1 To Me.ListView3.Columns.Count
                    l.SubItems.Add("")
                Next
                For Me.iterasi = 2 To tipeA

                    P_Stack = CDec(ListView2.Items(baris - 1).SubItems(2).Text) * CDec(ListView2.Items(baris - 1).SubItems(1).Text) * Val(TextBox6.Text) * Val(TextBox7.Text)

                    ListView3.Items(baris - 1).SubItems(1).Text = CDec((baris - 1) / 100)
                    ListView3.Items(baris - 1).SubItems(0).Text = baris


                    ListView3.Items(baris - 1).SubItems(2).Text = CDec(P_Stack)
                    If baris > 1 Then
                        ListView3.Items(baris - 1).SubItems(3).Text = ListView3.Items(baris - 2).SubItems(2).Text - ListView3.Items(baris - 1).SubItems(2).Text
                    Else
                        ListView3.Items(baris - 1).SubItems(3).Text = 0

                    End If


                    If ComboBox4.Text = "Real Time" Then

                        Chart2.Series(0).Points.AddXY(CDec(ListView3.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView3.Items(baris - 1).SubItems(2).Text))
                        Me.Chart2.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                        Me.Chart2.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                        If ComboBox1.Text = "Point" Then
                            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox1.Text = "Bar" Then
                            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox1.Text = "Area" Then
                            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox1.Text = "Fast Line" Then
                            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox3.Text = "Red" Then
                            Chart2.Series(0).Color = Color.Red
                        ElseIf ComboBox3.Text = "Green" Then
                            Chart2.Series(0).Color = Color.Green
                        ElseIf ComboBox3.Text = "Blue" Then
                            Chart2.Series(0).Color = Color.Blue
                        Else
                            Chart2.Series(0).Color = Color.Brown
                        End If
                        If ComboBox2.Text = "Dash" Then
                            With Chart2.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSe                                                                                                                                                                                                                                                                                                                                                                                                                                                    t
                            End With
                        Else
                            With Chart2.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)

                        Chart2.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
                        Chart2.ChartAreas("ChartArea1").AxisY.Title = "Power (Watts)"
                    End If
                Next
                TextBox32.Text = Format(CDbl(Me.Chart2.ChartAreas(0).AxisX.Maximum))
                TextBox36.Text = Format(CDbl(Me.Chart2.ChartAreas(0).AxisY.Maximum))
                TextBox96.Text = Format(CDbl(Me.Chart2.ChartAreas(0).AxisY.Minimum))
                TextBox95.Text = Format(CDbl(Me.Chart2.ChartAreas(0).AxisX.Minimum))
            Next

           

            Button28.PerformClick()
            Button45.PerformClick()
            Button23.PerformClick()
            MsgBox("finish")
           
        End If


    End Sub

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        Chart2.Series(0).Points.Clear()
        ' wait(1)

        For baris = 1 To ListView3.Items.Count

            'If ComboBox4.Text = "Real Time" Then

            Chart2.Series(0).Points.AddXY(ListView3.Items(baris - 1).SubItems(1).Text.ToString, (ListView3.Items(baris - 1).SubItems(2).Text))
            If ComboBox1.Text = "Point" Then
                Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
            ElseIf ComboBox1.Text = "Bar" Then
                Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
            ElseIf ComboBox1.Text = "Area" Then
                Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
            ElseIf ComboBox1.Text = "Fast Line" Then
                Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

            Else
                Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
            End If

            If ComboBox3.Text = "Red" Then
                Chart2.Series(0).Color = Color.Red
            ElseIf ComboBox3.Text = "Green" Then
                Chart2.Series(0).Color = Color.Green
            ElseIf ComboBox3.Text = "Blue" Then
                Chart2.Series(0).Color = Color.Blue
            Else
                Chart2.Series(0).Color = Color.Brown
            End If
            If ComboBox2.Text = "Dash" Then
                With Chart2.ChartAreas(0)
                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                End With
            Else
                With Chart2.ChartAreas(0)
                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                End With
            End If

            'wait(0.001)
            ' End If
            
        Next


        Chart2.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
        Chart2.ChartAreas("ChartArea1").AxisY.Title = "Power (Watts)"
        'e()
        ' Button14.PerformClick()
        ' Chart2.ChartAreas(0).AxisX.ScaleView.Size = [Double].NaN
        ' Chart2.ChartAreas(0).AxisY.ScaleView.Size = [Double].NaN
        'Chart2.ChartAreas(0).RecalculateAxesScale()
        ' chart2.BorderlineWidth = 0
        Me.Chart2.ChartAreas("ChartArea1").AxisY.IsStartedFromZero = False
        Me.Chart2.ChartAreas(0).AxisY.Maximum = (CDec(TextBox36.Text) + ((TextBox48.Text - TextBox47.Text) / 10))
        Me.Chart2.ChartAreas(0).AxisY.Minimum = (CDec(TextBox96.Text) - ((TextBox48.Text - TextBox47.Text) / 7))
        Me.Chart2.ChartAreas(0).AxisX.Maximum = (CDec(TextBox32.Text) * 100 + ((TextBox37.Text * 100 - 0) / 5))
        Me.Chart2.ChartAreas(0).AxisX.Minimum = (CDec(TextBox95.Text) * 100 - ((TextBox37.Text * 100 - 0) / 5))
        ' Me.Chart2.ChartAreas(0).AxisY.IsReversed = False
        Me.Chart2.ChartAreas(0).AxisX.IsStartedFromZero = False

        ' Me.Chart2.ChartAreas(0).AxisX.Minimum = TextBox25.Text - TextBox24.Text / 1000 - TextBox32.Text ' -0.001 'TextBox25.Text - TextBox37.Text / 10 - TextBox32.Text
        'Me.Chart2.ChartAreas("ChartArea1").AxisX.IsReversed = False
        Me.Chart2.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
        Me.Chart2.ChartAreas(0).AxisX.LabelStyle.Format = "N2"

        MsgBox("finish")
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        For g = 1 To ListView3.Items.Count
            If CDec(ListView3.Items(g - 1).SubItems(2).Text) = CDec(TextBox48.Text) Then
                'ListView3.Items(g - 1).BackColor = Color.LightGreen
                TextBox48.Text = Format(Val(TextBox48.Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = "                    max = " + Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                'chart2.Series(0).Points(g - 1).LabelFormat.ToLower()
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                Chart2.Series(0).Points(g - 1).LabelAngle = -60
            ElseIf ListView3.Items(g - 1).SubItems(3).Text = TextBox41.Text Then
                ' ListView3.Items(g - 1).BackColor = Color.Yellow

                TextBox39.BackColor = Color.Yellow
                TextBox39.Text = Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                'chart2.Series(0).Points(g).LabelForeColor = Color.Yellow


            ElseIf Val(ListView3.Items(g - 1).SubItems(3).Text) = Val(TextBox40.Text) Then
                ' ListView3.Items(g - 1).BackColor = Color.LightBlue
                ' TextBox40.BackColor = Color.LightBlue
                TextBox38.BackColor = Color.LightBlue
                TextBox38.Text = Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = "Peaky = " + Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                'chart2.Series(0).Points(g).LabelForeColor = Color.LightBlue
            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox47.Text Then
                'MsgBox(g)
                'MsgBox(ListView3.Items(g - 1).SubItems(2).Text)

                'ListView3.Items(g - 1).BackColor = Color.Orange
                ' TextBox47.BackColor = Color.Orange
                Chart2.Series(0).Points(g - 1).Label = "min = " + Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                ' MsgBox(chart2.Series(0).Points(g - 1).Label)
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.Orange
                'chart2.Series(0).Points(g).LabelForeColor = Color.Orange
            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox42.Text Then
                ' ListView3.Items(g - 1).BackColor = Color.LightCyan
                ' TextBox42.BackColor = Color.LightCyan
                Chart2.Series(0).Points(g - 1).Label = "mean = " + ListView3.Items(g - 1).SubItems(2).Text
                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 10
                Chart2.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                'chart2.Series(0).Points(g).LabelForeColor = Color.LightCyan
            End If

        Next g
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        TextBox48.Text = ListView3.Items(0).SubItems(2).Text
        Dim g As Integer
        'Dim 1 as index
        For g = 2 To ListView3.Items.Count

            If CDec(ListView3.Items(g - 2).SubItems(2).Text) < CDec(ListView3.Items(g - 1).SubItems(2).Text) And TextBox48.Text < CDec(ListView3.Items(g - 1).SubItems(2).Text) Then
                TextBox48.Text = CDec(ListView3.Items(g - 1).SubItems(2).Text)
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)

            Else
                TextBox48.Text = TextBox48.Text
                ''  End If

                ' End If

            End If

        Next

        TextBox47.Text = ListView3.Items(0).SubItems(2).Text
        Dim i As Integer
        For i = 2 To ListView3.Items.Count

            'jika angka pertama lebih kecil dari pada angka kedua dan textbox 34 lebih besar dari pada angka kedua maka
            'nilai terkecil adalah di TextBox47
            If CDec(ListView3.Items(i - 1).SubItems(2).Text) < CDec(ListView3.Items(i - 2).SubItems(2).Text) And TextBox47.Text > CDec(ListView3.Items(i - 1).SubItems(2).Text) Then
                TextBox47.Text = CDec(ListView3.Items(i - 1).SubItems(2).Text)
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)

                ' ListView3.Items(i - 2).BackColor = Color.Red
            Else
                TextBox47.Text = Val(TextBox47.Text)
                ''  End If

                ' End If

            End If

        Next i
        TextBox46.Text = TextBox48.Text - TextBox47.Text
        TextBox37.Text = ListView3.Items(q - 1).SubItems(1).Text


        TextBox40.Text = Math.Abs(CDec(ListView3.Items(0).SubItems(3).Text))

        'Dim 1 as index
        For g = 2 To ListView3.Items.Count

            If Math.Abs(CDec(ListView3.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(ListView3.Items(g - 1).SubItems(3).Text)) And TextBox40.Text < Math.Abs(CDec(ListView3.Items(g - 1).SubItems(3).Text)) Then
                TextBox40.Text = Math.Abs(CDec(ListView3.Items(g - 1).SubItems(3).Text))
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)
                ' ListView3.Items(g - 1).BackColor = Color.Blue
                ' ListView3.Items(g - 2).BackColor = Color.Blue
                
            Else
                TextBox40.Text = Val(TextBox40.Text)

                ''  End If

                ' End If

            End If

        Next g


        TextBox41.Text = Math.Abs(CDec(ListView3.Items(1).SubItems(3).Text))
        'Dim i As Integer

        For i = 3 To ListView3.Items.Count

            'jika angka pertama lebih kecil dari pada angka kedua dan textbox 7 lebih besar dari pada angka kedua maka
            'nilai terkecil adalah di textbox7
            If Math.Abs(CDec(ListView3.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(ListView3.Items(i - 2).SubItems(3).Text)) And TextBox41.Text > Math.Abs(CDec(ListView3.Items(i - 1).SubItems(3).Text)) Then
                TextBox41.Text = Math.Abs(CDec(ListView3.Items(i - 2).SubItems(3).Text))
                '  ElseIf CDec(listview1.Items(i - 1).SubItems(1).Text) < CDec(listview1.Items(i - 2).SubItems(1).Text) And TextBox7.Text > CDec(listview1.Items(i - 1).SubItems(1).Text) Then
                '   TextBox7.Text = CDec(listview1.Items(i - 2).SubItems(1).Text)
                ' ListView3.Items(i - 1).BackColor = Color.Yellow
                ' ListView3.Items(i - 2).BackColor = Color.Yellow
                ' If ListView3.Items(i - 1).SubItems(3).Text < 0 Then
                'TextBox41.Text = Val("-" + TextBox41.Text)

                ' End If
            Else
                ' If ListView3.Items(i - 1).SubItems(3).Text < 0 Then
                ' TextBox41.Text = Val("-" + TextBox41.Text)
                '   Else
                TextBox41.Text = Val(TextBox41.Text)
                ' End If

                ''  End If

                ' End If

            End If

        Next i

        'Dim maximum As String
        'Dim minimum As String
        ' maximum = (TextBox48.Text + (TextBox36.Text))
        '  minimum = (TextBox40.Text - (TextBox36.Text))
        ' Dim mini As String

        ' If baris > 2 Then

        'End If
        '  chart2.ChartAreas(0).AxisY.Minimum = TextBox48.Text 'Then
        '  chart2.ChartAreas(0).AxisY.Maximum = TextBox47.Text
        'chart2.AutoScaling = True
        'chart2.Series.Dispose()

        Chart2.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
        Chart2.ChartAreas("ChartArea1").AxisY.Title = "Power (Watts)"
        ' chart2.ChartAreas(0).AxisX.ScaleView.Size = Val(baris)
        ' chart2.ChartAreas(0).AxisY.ScaleView.Size = Val(TextBox46.Text)
        wait(0.001)
        TextBox42.Text = "0"
        For i = 1 To ListView3.Items.Count
            TextBox42.Text = TextBox42.Text + CDec(ListView3.Items(i - 1).SubItems(2).Text)
        Next i
        TextBox42.Text = TextBox42.Text / ListView3.Items.Count
        For g = 1 To ListView3.Items.Count
            If CDec(ListView3.Items(g - 1).SubItems(2).Text) = CDec(TextBox48.Text) Then
                ListView3.Items(g - 1).BackColor = Color.LightGreen
                TextBox48.BackColor = Color.LightGreen


            ElseIf ListView3.Items(g - 1).SubItems(3).Text = TextBox41.Text Then
                ListView3.Items(g - 1).BackColor = Color.Yellow
                TextBox41.BackColor = Color.Yellow
                TextBox39.Text = ListView3.Items(g - 1).SubItems(2).Text
                TextBox39.BackColor = Color.Yellow

            ElseIf Val(ListView3.Items(g - 1).SubItems(3).Text) = Val(TextBox40.Text) Then
                ListView3.Items(g - 1).BackColor = Color.LightBlue
                TextBox40.BackColor = Color.LightBlue
                TextBox38.Text = ListView3.Items(g - 1).SubItems(2).Text
                TextBox38.BackColor = Color.LightBlue
            ElseIf ListView3.Items(g - 1).SubItems(3).Text = "-" & TextBox40.Text Then
                ListView3.Items(g - 1).BackColor = Color.LightBlue
                TextBox40.BackColor = Color.LightBlue
                TextBox38.Text = ListView3.Items(g - 1).SubItems(2).Text
                TextBox38.BackColor = Color.LightBlue
                TextBox40.Text = "-" & TextBox40.Text
            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox47.Text Then


                ListView3.Items(g - 1).BackColor = Color.Orange
                TextBox47.BackColor = Color.Orange

            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox42.Text Then
                ListView3.Items(g - 1).BackColor = Color.LightCyan
                TextBox42.BackColor = Color.LightCyan

            End If
            
        Next g
        ' Button45.PerformClick()
        'Chart2.ChartAreas(0).AxisY.Interval = TextBox46.Text / 10
        ' Chart2.ChartAreas(0).AxisX.Interval = TextBox24.Text / 10

        'chart2.ChartAreas(0).AxisY2.LineWidth = 0
        'chart2.ChartAreas(0).BorderDashStyle.DashDot = True
        'chart2.ChartAreas(0).BackImageAlignment.linewidth = 0
        ' chart2.ChartAreas(0).AxisY.LabelStyle.Enabled = False
        ' Button23.PerformClick()
    End Sub

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button47.Click
        Dim openFileDialog1 As New OpenFileDialog()
        'set the root to the z drive
        openFileDialog1.InitialDirectory = "D:\"
        'make sure the root goes back to where the user started
        ''   openFileDialog1.RestoreDirectory = True
        'show the dialog
        ''  openFileDialog1.ShowDialog()
        ''
        '' If (openFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then
        ''Path = openFileDialog1.FileName
        ''  End If

        ' Call ShowDialog.
        Dim result As DialogResult = openFileDialog1.ShowDialog()

        ' Test result.
        Dim path As String = openFileDialog1.FileName
        If result = Windows.Forms.DialogResult.OK Then
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=Excel 12.0;"
            ' Get the file name.

            '  Try
            Dim table As DataTable = New DataTable()
            Dim excelName As String = "Sheet1"
            Dim strConnection As String = String.Format(connStr)
            Dim conn As OleDbConnection = New OleDbConnection(strConnection)
            conn.Open()
            Dim oada As OleDbDataAdapter = New OleDbDataAdapter("select * from [" & excelName & "$]", strConnection)
            table.TableName = "TableInfo"
            oada.Fill(table)
            conn.Close()
            ListView3.Items.Clear()

            For i As Integer = 0 To table.Rows.Count - 1
                Dim drow As DataRow = table.Rows(i)

                If drow.RowState <> DataRowState.Deleted Then
                    Dim lvi As ListViewItem = New ListViewItem(drow("No").ToString())
                    lvi.SubItems.Add(drow("Curr density (A/cm^2)").ToString())
                    lvi.SubItems.Add(drow("Power (Watts)").ToString())
                    lvi.SubItems.Add(drow("Sequence Discreapancy").ToString())
                    ListView3.Items.Add(lvi)
                End If
            Next
            ' Read in text.
            '' Dim text As String = File.ReadAllText(path)

            ' For debugging.
            ''  Me.Text = text.Length.ToString

            '  Catch ex As Exception

            ' Report an error.
            ''   Me.Text = "Error"

            ' End Try
        End If


        ' For Me.baris = 1 To Me.listview3.Items.Count
        ' l = Me.listview3.Items.Add("")
        ' For j As Integer = 1 To Me.listview3.Columns.Count
        ' l.SubItems.Add("")
        'Next
        'listview3.Items(baris - 1).SubItems(0).Text = baris
        ' Next


    End Sub



    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button48.Click

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView3.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView3.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView3.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView3.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView3.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView3.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        ListView3.Items.Clear()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        For g = 1 To ListView3.Items.Count
            If CDec(ListView3.Items(g - 1).SubItems(2).Text) = CDec(TextBox48.Text) Then
                'ListView3.Items(g - 1).BackColor = Color.LightGreen
                TextBox48.Text = Format(Val(TextBox48.Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = ""
                 Chart2.Series(0).Points(g - 1).MarkerSize = 0
              
            ElseIf ListView3.Items(g - 1).SubItems(3).Text = TextBox41.Text Then
                ' ListView3.Items(g - 1).BackColor = Color.Yellow

                TextBox39.BackColor = Color.Yellow
                TextBox39.Text = Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = ""
                Chart2.Series(0).Points(g - 1).MarkerSize = 0
               


            ElseIf Val(ListView3.Items(g - 1).SubItems(3).Text) = Val(TextBox40.Text) Then
                
                TextBox38.BackColor = Color.LightBlue
                TextBox38.Text = Format(Val(ListView3.Items(g - 1).SubItems(2).Text), "0.000")
                Chart2.Series(0).Points(g - 1).Label = ""
                Chart2.Series(0).Points(g - 1).MarkerSize = 0
               
            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox47.Text Then
               
                Chart2.Series(0).Points(g - 1).Label = ""

                Chart2.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart2.Series(0).Points(g - 1).MarkerSize = 0
             
            ElseIf ListView3.Items(g - 1).SubItems(2).Text = TextBox42.Text Then
        
                Chart2.Series(0).Points(g - 1).Label = ""

                Chart2.Series(0).Points(g - 1).MarkerSize = 0
       
            End If

        Next g
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        With Chart2.ChartAreas("ChartArea1")
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
           
        End With
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot

        End With
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim saveFileDialog1 As New SaveFileDialog()

        ' Sets the current file name filter string, which determines 
        ' the choices that appear in the "Save as file type" or 
        ' "Files of type" box in the dialog box.
        saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        ' Set image file format
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As ChartImageFormat = ChartImageFormat.Bmp

            If saveFileDialog1.FileName.EndsWith("bmp") Then
                format = ChartImageFormat.Bmp
            Else
                If saveFileDialog1.FileName.EndsWith("jpg") Then
                    format = ChartImageFormat.Jpeg
                Else
                    If saveFileDialog1.FileName.EndsWith("emf") Then
                        format = ChartImageFormat.Emf
                    Else
                        If saveFileDialog1.FileName.EndsWith("gif") Then
                            format = ChartImageFormat.Gif
                        Else
                            If saveFileDialog1.FileName.EndsWith("png") Then
                                format = ChartImageFormat.Png
                            Else
                                If saveFileDialog1.FileName.EndsWith("tif") Then
                                    format = ChartImageFormat.Tiff

                                End If
                            End If ' Save image
                        End If
                    End If
                End If
            End If
            Chart2.SaveImage(saveFileDialog1.FileName, format)
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Chart2.Series(0).Points.Clear()

    End Sub

    Private Sub Button17_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
          
        End With
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        TextBox39.Text = ""
        TextBox39.BackColor = Color.White
        TextBox40.Text = ""
        TextBox40.BackColor = Color.White
        TextBox41.Text = ""
        TextBox41.BackColor = Color.White
        TextBox42.Text = ""
        TextBox42.BackColor = Color.White
        TextBox46.Text = ""
        TextBox46.BackColor = Color.White
        TextBox47.Text = ""
        TextBox47.BackColor = Color.White
        TextBox48.Text = ""
        TextBox48.BackColor = Color.White
        TextBox37.Text = ""
        TextBox37.BackColor = Color.White
        TextBox38.Text = ""
        TextBox38.BackColor = Color.White
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Button21.PerformClick()
        Button27.PerformClick()
        Button46.PerformClick()

    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button44.Click
        Chart2.ChartAreas("ChartArea1").AxisX.ScaleView.Size = TextBox24.Text * 2
        Chart2.ChartAreas("ChartArea1").AxisY.ScaleView.Size = TextBox46.Text * 2
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        Chart2.ChartAreas("ChartArea1").AxisX.ScaleView.Size = TextBox24.Text / 2
        Chart2.ChartAreas("ChartArea1").AxisY.ScaleView.Size = TextBox46.Text / 2
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
        End With
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
        End With
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        With Chart2.ChartAreas("ChartArea1")
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button57.Click
        Dim N_cells As String = TextBox7.Text
        Dim A_cell As String = TextBox6.Text
        Dim Hyd_usage As String
        Dim F As String = TextBox2.Text
        'N = ListView2.Items.Count
        z = TextBox25.Text
        
        For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01
            l = Me.ListView4.Items.Add("")
            For j As Integer = 1 To Me.ListView3.Columns.Count
                l.SubItems.Add("")
            Next

            P_Stack = N_cells * (CDec(ListView2.Items(baris - 1).SubItems(2).Text) * (CDec(ListView3.Items(baris - 1).SubItems(1).Text) * A_cell))
            If CDec(ListView2.Items(baris - 1).SubItems(2).Text) < Val(0) Then
                Hyd_usage = 0
            Else
                Hyd_usage = (P_Stack) / (2 * CDec(ListView2.Items(baris - 1).SubItems(2).Text) * F)
            End If

            ListView4.Items(baris - 1).SubItems(0).Text = baris
            ListView4.Items(baris - 1).SubItems(1).Text = CDec(P_Stack)
            ListView4.Items(baris - 1).SubItems(2).Text = CDec(Hyd_usage)
            If Not baris < 2 Then
                ListView4.Items(baris - 1).SubItems(3).Text = CDec(ListView4.Items(baris - 1).SubItems(2).Text) - CDec(ListView4.Items(baris - 2).SubItems(2).Text)

            End If
            If ComboBox8.Text = "Real Time" Then
                Me.Chart3.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                Me.Chart3.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                Chart3.Series(0).Points.AddXY(CDec(ListView4.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView4.Items(baris - 1).SubItems(2).Text))
                Me.Chart3.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
                Me.Chart3.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
                If ComboBox5.Text = "Point" Then
                    Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                ElseIf ComboBox5.Text = "Bar" Then
                    Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                ElseIf ComboBox5.Text = "Area" Then
                    Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                ElseIf ComboBox5.Text = "Fast Line" Then
                    Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                Else
                    Chart3.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                End If

                If ComboBox7.Text = "Red" Then
                    Chart3.Series(0).Color = Color.Red
                ElseIf ComboBox7.Text = "Green" Then
                    Chart3.Series(0).Color = Color.Green
                ElseIf ComboBox7.Text = "Blue" Then
                    Chart3.Series(0).Color = Color.Blue
                Else
                    Chart3.Series(0).Color = Color.Brown
                End If
                If ComboBox6.Text = "Dash" Then
                    With Chart3.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSe                                                                                                                                                                                                                                                                                                                                                                                                                                                    t
                    End With
                Else
                    With Chart3.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                End If

                wait(0.001)

                Chart3.ChartAreas("ChartArea1").AxisY.Title = "P_Stack (Watts)"
                Chart3.ChartAreas("ChartArea1").AxisX.Title = "Hyd_usage (m^3/s)"
               
            End If


        Next
        MsgBox("Finish")
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        'P_Stack = N_cells * (V_out * (i * A_cell))
        Dim Water_prod As String
        Dim mm_H2o As String = TextBox19.Text
        Dim F As String = TextBox2.Text

       

        For Me.baris = 1 To ListView3.Items.Count 'Step 0.01
            l = Me.ListView5.Items.Add("")
            For j As Integer = 1 To Me.ListView5.Columns.Count
                l.SubItems.Add("")
            Next
            If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                Water_prod = 0
            Else
                Water_prod = (mm_H2o * ListView3.Items(baris - 1).SubItems(2).Text) / (3600 * ListView2.Items(baris - 1).SubItems(2).Text * F)
            End If
           
            ListView5.Items(baris - 1).SubItems(0).Text = baris
            ListView5.Items(baris - 1).SubItems(1).Text = CDec(ListView2.Items(baris - 1).SubItems(2).Text)
            ListView5.Items(baris - 1).SubItems(2).Text = CDec(Water_prod)
            If Not baris < 2 Then
                ListView5.Items(baris - 1).SubItems(3).Text = Val(ListView5.Items(baris - 1).SubItems(2).Text) - Val(ListView5.Items(baris - 2).SubItems(2).Text)

            End If
            If ComboBox15.Text = "Real Time" Then

                Chart4.Series(0).Points.AddXY((ListView5.Items(baris - 1).SubItems(1).Text.ToString), (ListView5.Items(baris - 1).SubItems(2).Text))
                Me.Chart4.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                Me.Chart4.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                If ComboBox9.Text = "Point" Then
                    Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                ElseIf ComboBox9.Text = "Bar" Then
                    Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                ElseIf ComboBox9.Text = "Area" Then
                    Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                ElseIf ComboBox9.Text = "Fast Line" Then
                    Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                Else
                    Chart4.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                End If

                If ComboBox11.Text = "Red" Then
                    Chart4.Series(0).Color = Color.Red
                ElseIf ComboBox11.Text = "Green" Then
                    Chart4.Series(0).Color = Color.Green
                ElseIf ComboBox11.Text = "Blue" Then
                    Chart4.Series(0).Color = Color.Blue
                Else
                    Chart4.Series(0).Color = Color.Brown
                End If
                If ComboBox10.Text = "Dash" Then
                    With Chart3.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSe                                                                                                                                                                                                                                                                                                                                                                                                                                                    t
                    End With
                Else
                    With Chart3.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                End If

                wait(0.001)

                Chart4.ChartAreas("ChartArea1").AxisX.Title = "P_Stack (Watts)"
                Chart4.ChartAreas("ChartArea1").AxisY.Title = "Water Produced"
            End If
            Me.Chart4.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
            Me.Chart4.ChartAreas(0).AxisX.LabelStyle.Format = "N2"

        Next
    End Sub

    Private Sub Button141_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button141.Click
        Dim N_cells As String = TextBox7.Text
        Dim A_cell As String = TextBox6.Text
        Dim Hyd_usage As String
        Dim F As String = TextBox2.Text
        'N = ListView2.Items.Count
        z = TextBox25.Text

        For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01
            l = Me.ListView8.Items.Add("")
            For j As Integer = 1 To Me.ListView3.Columns.Count
                l.SubItems.Add("")
            Next

            ListView8.Items(baris - 1).SubItems(0).Text = baris
            ListView8.Items(baris - 1).SubItems(2).Text = ListView3.Items(baris - 1).SubItems(2).Text ' CDec(P_Stack)
            ListView8.Items(baris - 1).SubItems(1).Text = ListView4.Items(baris - 1).SubItems(2).Text * 22.4 / 1000 'CDec(Hyd_usage)
            If Not baris < 2 Then
                ListView8.Items(baris - 1).SubItems(3).Text = CDec(ListView8.Items(baris - 1).SubItems(2).Text) - CDec(ListView8.Items(baris - 2).SubItems(2).Text)

            End If
            If ComboBox8.Text = "Real Time" Then
               
                Chart7.Series(0).Points.AddXY(CDec(ListView8.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView8.Items(baris - 1).SubItems(2).Text))
                Me.Chart7.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
                Me.Chart7.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
                If ComboBox5.Text = "Point" Then
                    Chart7.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                ElseIf ComboBox5.Text = "Bar" Then
                    Chart7.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                ElseIf ComboBox5.Text = "Area" Then
                    Chart7.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                ElseIf ComboBox5.Text = "Fast Line" Then
                    Chart7.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                Else
                    Chart7.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                End If

                If ComboBox7.Text = "Red" Then
                    Chart7.Series(0).Color = Color.Red
                ElseIf ComboBox7.Text = "Green" Then
                    Chart7.Series(0).Color = Color.Green
                ElseIf ComboBox7.Text = "Blue" Then
                    Chart7.Series(0).Color = Color.Blue
                Else
                    Chart7.Series(0).Color = Color.Brown
                End If
                If ComboBox6.Text = "Dash" Then
                    With Chart7.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSe                                                                                                                                                                                                                                                                                                                                                                                                                                                    t
                    End With
                Else
                    With Chart7.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                End If
                Me.Chart7.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                Me.Chart7.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                wait(0.001)

                Chart7.ChartAreas("ChartArea1").AxisY.Title = "P_Stack (Watts)"
                Chart7.ChartAreas("ChartArea1").AxisX.Title = "Hyd_usage (m^3/s)"

            End If


        Next
        Me.Chart7.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
        Me.Chart7.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
        MsgBox("Finish")
    End Sub
   
    Private Sub Button150_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button150.Click

    End Sub

    Private Sub Button153_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button153.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView8.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView8.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView8.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView8.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView8.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView8.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button69.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView4.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView4.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView4.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView4.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView4.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView4.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button162_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button162.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325

        Dim i As String
        Dim P_air As String = TextBox5.Text * 101325

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text


        Chart8.Series(0).Points.Clear()
        ListView8.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 


        z = TextBox25.Text

        For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

            l = Me.ListView9.Items.Add("")
            For j As Integer = 1 To Me.ListView9.Columns.Count
                l.SubItems.Add("")
            Next
            Try
                For Me.iterasi = 2 To tipeA

                    ListView9.Items(baris - 1).SubItems(0).Text = baris

                    i = CStr((baris - 1) * TextBox26.Text) '* Math.Sqrt(2)))
                    'Calculation of Partial Pressures 
                    'Calculation of saturation pressure of water 

                    x = -2.1794 + (0.02953 * CDec(Tc)) - CDec(9.1837 * (10 ^ -5) * (Tc ^ 2)) + (1.4454 * (10 ^ -7) * (CDec(Tc) ^ 3))
                    P_H2O = (10 ^ x) * 101325
                    'Calculation of partial pressure of hydrogen 
                    pp_H2 = (0.5 * CDec((P_H2) / (Math.Exp(1.653 * i / (tk ^ 1.334))) - P_H2O))
                   
                    ListView9.Items(baris - 1).SubItems(1).Text = CDec(pp_H2)
                    ListView9.Items(baris - 1).SubItems(2).Text = ListView3.Items(baris - 1).SubItems(2).Text
                        If ComboBox12.Text = "Real Time" Then

                        Chart8.Series(0).Points.AddXY(CDec(ListView9.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView9.Items(baris - 1).SubItems(2).Text))
                       
                        If ComboBox17.Text = "Point" Then
                            Chart8.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart8.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart8.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart8.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart8.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                            If ComboBox13.Text = "Red" Then
                            Chart8.Series(0).Color = Color.Red
                            ElseIf ComboBox13.Text = "Green" Then
                            Chart8.Series(0).Color = Color.Green
                            ElseIf ComboBox13.Text = "Blue" Then
                            Chart8.Series(0).Color = Color.Blue
                            Else
                            Chart8.Series(0).Color = Color.Brown
                            End If
                            If ComboBox14.Text = "Dash" Then
                            With Chart8.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                            Else
                            With Chart8.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                            End If

                            wait(0.001)
                    End If
                   
                    Chart8.ChartAreas("ChartArea1").AxisX.Title = "Partial Pressure Hydrogen (Pa)"
                    Chart8.ChartAreas("ChartArea1").AxisY.Title = "Power (watt)"
                   
                Next

            Catch t As Exception
            End Try
        Next
        
        Me.Chart8.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
        Me.Chart8.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
        
        MsgBox("finish")

    End Sub

    Private Sub Button174_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button174.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView9.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView9.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView9.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView9.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView9.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView9.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button99_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button99.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325

        Dim i As String
        Dim P_air As String = TextBox5.Text * 101325

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text
        Dim H_G As String
        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text
        Dim E_HHV As String = TextBox22.Text
        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text


        Chart5.Series(0).Points.Clear()
        ListView6.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 


        z = TextBox25.Text

        For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

            l = Me.ListView6.Items.Add("")
            For j As Integer = 1 To Me.ListView6.Columns.Count
                l.SubItems.Add("")
            Next
            Try
                For Me.iterasi = 2 To tipeA

                    ListView6.Items(baris - 1).SubItems(0).Text = baris
                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        H_G = 0
                    Else
                        H_G = CDec(ListView2.Items(baris - 1).SubItems(1).Text * TextBox7.Text * (E_HHV - ListView2.Items(baris - 1).SubItems(2).Text))
                    End If

                    ListView6.Items(baris - 1).SubItems(1).Text = CDec(ListView3.Items(baris - 1).SubItems(2).Text)
                    ListView6.Items(baris - 1).SubItems(2).Text = CDec(H_G)
                    If ComboBox12.Text = "Real Time" Then

                        Chart5.Series(0).Points.AddXY(CDec(ListView6.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView6.Items(baris - 1).SubItems(2).Text))
                        Me.Chart5.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                        Me.Chart5.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                        If ComboBox17.Text = "Point" Then
                            Chart5.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart5.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart5.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart5.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart5.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox13.Text = "Red" Then
                            Chart5.Series(0).Color = Color.Red
                        ElseIf ComboBox13.Text = "Green" Then
                            Chart5.Series(0).Color = Color.Green
                        ElseIf ComboBox13.Text = "Blue" Then
                            Chart5.Series(0).Color = Color.Blue
                        Else
                            Chart5.Series(0).Color = Color.Brown
                        End If
                        If ComboBox14.Text = "Dash" Then
                            With Chart5.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        Else
                            With Chart5.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)
                    End If
                    Me.Chart5.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
                    Me.Chart5.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
                    Chart5.ChartAreas("ChartArea1").AxisX.Title = "Partial Pressure Hydrogen (Pa)"
                    Chart5.ChartAreas("ChartArea1").AxisY.Title = "Power (watt)"

                Next

            Catch t As Exception
            End Try
        Next



        MsgBox("finish")

    End Sub

    Private Sub Button120_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button120.Click

        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325
        Dim Eff_HHV As String
        Dim UF As String = TextBox16.Text
        Dim E_HHV As String = TextBox22.Text
        Dim i As String
        Dim P_air As String = TextBox5.Text * 101325

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text

        Chart6.Series(0).Points.Clear()
        ListView7.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 


        z = TextBox25.Text

        For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

            l = Me.ListView7.Items.Add("")
            For j As Integer = 1 To Me.ListView7.Columns.Count
                l.SubItems.Add("")
            Next
            Try
                For Me.iterasi = 2 To tipeA

                    ListView7.Items(baris - 1).SubItems(0).Text = baris

                  

                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        Eff_HHV = 0
                    Else
                        Eff_HHV = (UF * ListView2.Items(baris - 1).SubItems(2).Text * 100) / (E_HHV)
                    End If
                    ListView7.Items(baris - 1).SubItems(1).Text = CDec(ListView2.Items(baris - 1).SubItems(1).Text)
                    ListView7.Items(baris - 1).SubItems(2).Text = CDec(Eff_HHV)
                    If ComboBox12.Text = "Real Time" Then

                        Chart6.Series(0).Points.AddXY(CDec(ListView7.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView7.Items(baris - 1).SubItems(2).Text))
                        Me.Chart6.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "N2"
                        Me.Chart6.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "N2"
                        If ComboBox17.Text = "Point" Then
                            Chart6.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart6.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart6.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart6.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart6.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox13.Text = "Red" Then
                            Chart6.Series(0).Color = Color.Red
                        ElseIf ComboBox13.Text = "Green" Then
                            Chart6.Series(0).Color = Color.Green
                        ElseIf ComboBox13.Text = "Blue" Then
                            Chart6.Series(0).Color = Color.Blue
                        Else
                            Chart6.Series(0).Color = Color.Brown
                        End If
                        If ComboBox14.Text = "Dash" Then
                            With Chart6.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        Else
                            With Chart6.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)
                    End If
                    Me.Chart6.ChartAreas(0).AxisY.LabelStyle.Format = "N2"
                    Me.Chart6.ChartAreas(0).AxisX.LabelStyle.Format = "N2"
                    Chart6.ChartAreas("ChartArea1").AxisX.Title = "Partial Pressure Hydrogen (Pa)"
                    Chart6.ChartAreas("ChartArea1").AxisY.Title = "Power (watt)"

                Next

            Catch t As Exception
            End Try
        Next



        MsgBox("finish")
    End Sub
End Class
