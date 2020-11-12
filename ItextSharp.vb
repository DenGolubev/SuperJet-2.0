Imports System.IO
Imports System.Data
Imports System.Reflection
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Module ItextSharp

    Sub New_PDF(DGV As DataGridView)
        Dim pdfDoc As New Document()
        Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(pdfDoc, New FileStream("Отчет - " & Форма_Конвертация.ComboBox1.Text & ".pdf", FileMode.Create))

        pdfDoc.Open()

        Dim f As Font = GetFont_14()
        Dim f1 As Font = GetFont_18()

        pdfDoc.Add(New Paragraph("Отчет по СКС " & " " & Форма_Конвертация.ComboBox1.Text & " " & "C -" & Форма_Конвертация.DateTimePicker1.Value.ToShortDateString &
                                 " " & "По -" & Форма_Конвертация.DateTimePicker2.Value.ToShortDateString, f1))
        pdfDoc.Add(New Paragraph("-----------------------------------------------------------------------------------------------------------", f))
        pdfDoc.Add(New Paragraph("Отчет по первичным подключениям", f1))
        'pdfDoc.Add(New Paragraph(Форма_Конвертация.Label2.Text & "" & " - " & Форма_Конвертация.TextBox2.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label16.Text & "" & " - " & Форма_Конвертация.TextBox16.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label3.Text & "" & " - " & Форма_Конвертация.TextBox3.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label4.Text & "" & " - " & Форма_Конвертация.TextBox4.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label5.Text & "" & " - " & Форма_Конвертация.TextBox5.Text, f))

        pdfDoc.Add(New Paragraph("-----------------------------------------------------------------------------------------------------------", f))
        pdfDoc.Add(New Paragraph("Статистика по услугам", f1))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label18.Text & "" & " - " & Форма_Конвертация.TextBox18.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label19.Text & "" & " - " & Форма_Конвертация.TextBox19.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label8.Text & "" & " - " & Форма_Конвертация.TextBox8.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label9.Text & "" & " - " & Форма_Конвертация.TextBox9.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label13.Text & "" & " - " & Форма_Конвертация.TextBox13.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label12.Text & "" & " - " & Форма_Конвертация.TextBox12.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label15.Text & "" & " - " & Форма_Конвертация.TextBox15.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label17.Text & "" & " - " & Форма_Конвертация.TextBox17.Text, f))

        pdfDoc.Add(New Paragraph("-----------------------------------------------------------------------------------------------------------", f))
        pdfDoc.Add(New Paragraph("Статистика по ТТ", f1))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label10.Text & "" & " - " & Форма_Конвертация.TextBox10.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label11.Text & "" & " - " & Форма_Конвертация.TextBox11.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label22.Text & "" & " - " & Форма_Конвертация.TextBox21.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label23.Text & "" & " - " & Форма_Конвертация.TextBox23.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label24.Text & "" & " - " & Форма_Конвертация.TextBox24.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label25.Text & "" & " - " & Форма_Конвертация.TextBox26.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label27.Text & "" & " - " & Форма_Конвертация.TextBox28.Text, f))

        pdfDoc.Add(New Paragraph("-----------------------------------------------------------------------------------------------------------", f))
        pdfDoc.Add(New Paragraph("Статистика по СКС", f1))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label20.Text & "" & " - " & Форма_Конвертация.TextBox20.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label6.Text & "" & " - " & Форма_Конвертация.TextBox6.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label30.Text & "" & " - " & Форма_Конвертация.TextBox30.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label7.Text & "" & " - " & Форма_Конвертация.TextBox7.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label14.Text & "" & " - " & Форма_Конвертация.TextBox14.Text, f))
        pdfDoc.Add(New Paragraph(Форма_Конвертация.Label29.Text & "" & " - " & Форма_Конвертация.TextBox29.Text, f))


        pdfDoc.Close()
    End Sub

    Private Function GetFont_14() As Font
        Dim fontLocation As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF")

        Dim baseFont As BaseFont = BaseFont.CreateFont(fontLocation, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED)

        Return New Font(baseFont, 14, 0)
    End Function

    Private Function GetFont_18() As Font
        Dim fontLocation As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF")

        Dim baseFont As BaseFont = BaseFont.CreateFont(fontLocation, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED)

        Return New Font(baseFont, 18, 0)
    End Function

    Sub DGV_PDF(DGV As DataGridView, C_box As ComboBox)
        Dim f As Font = GetFont_14()
        Dim f1 As Font = GetFont_18()
        Dim pdfTable As New PdfPTable(DGV.ColumnCount)

        'Add Title of PDF
        Dim cell1 As New PdfPCell(New Phrase(C_box.Text, f1))
        cell1.Colspan = 28
        cell1.Border = 0
        cell1.HorizontalAlignment = 1
        pdfTable.AddCell(cell1)
        'pdfDoc.Add(New Paragraph("-----------------------------------------------------------------------------------------------------------", f))

        'Creating iTextSharp Table from the DataTable data
        pdfTable.DefaultCell.Padding = 3
        pdfTable.WidthPercentage = 100
        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
        pdfTable.DefaultCell.BorderWidth = 1


        'Adding Header row
        For Each column As DataGridViewColumn In DGV.Columns
            Dim cell As New PdfPCell(New Phrase(column.HeaderText, f))
            'cell.BackgroundColor = New ItextSharp.text.BaseColor(240, 240, 240)
            pdfTable.AddCell(cell)
        Next

        'Adding DataRow
        For Each row As DataGridViewRow In DGV.Rows
            If row.IsNewRow = False Then
                For Each cell As DataGridViewCell In row.Cells
                    pdfTable.AddCell(cell.Value.ToString())
                Next
            End If
        Next

        'Exporting to PDF
        Dim folderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Отчеты\"
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
        End If
        Using stream As New FileStream(folderPath & C_box.Text & ".pdf", FileMode.Create)
            Dim fontLocation As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF")
            Dim baseFont As BaseFont = BaseFont.CreateFont(fontLocation, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED)
            Dim pdfDoc As New Document(PageSize.A4.Rotate, 5.0F, 5.0F, 5.0F, 5.0F)
            PdfWriter.GetInstance(pdfDoc, stream)
            pdfDoc.Open()
            pdfDoc.Add(pdfTable)
            pdfDoc.Close()
            stream.Close()
        End Using

        MsgBox("PDF Created")
    End Sub


End Module
