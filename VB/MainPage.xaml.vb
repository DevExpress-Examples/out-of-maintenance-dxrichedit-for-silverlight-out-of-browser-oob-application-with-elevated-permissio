Option Strict Off

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Runtime.InteropServices.Automation

Namespace RichEditOOBElevatedPermissions
	Partial Public Class MainPage
		Inherits UserControl
		Private featureComplete As Boolean = Application.Current.HasElevatedPermissions
		Private Const errorMessage As String = "This application is not trusted."
        Private outlook As Object

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnLoadImage_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If featureComplete Then
                richEditControl1.Document.Images.Insert(richEditControl1.Document.Range.End, DocumentImageSource.FromUri("http://www.devexpress.com/Home/i/logos/preview.png", richEditControl1))
			Else
				MessageBox.Show(errorMessage)
			End If
		End Sub

		Private Sub btnLoad_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If featureComplete Then
				Dim myDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
				Dim fileName As String = "test.rtf"
                Dim path As String = System.IO.Path.Combine(myDocuments, fileName)

                If File.Exists(path) Then
                    richEditControl1.RtfText = File.ReadAllText(path)
                Else
                    MessageBox.Show(String.Format("The '{0}' file does not exist.", path))
                End If
            Else
                MessageBox.Show(errorMessage)
            End If
		End Sub

		Private Sub btnSave_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If featureComplete Then
				Dim myDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
				Dim fileName As String = "test.rtf"
                Dim path As String = System.IO.Path.Combine(myDocuments, fileName)

				File.WriteAllText(path, richEditControl1.RtfText)
			Else
				MessageBox.Show(errorMessage)
			End If
		End Sub

		Private Sub btnEmail_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If AutomationFactory.IsAvailable Then
				If InitializeOutlook() Then
                    Dim mailItem = outlook.CreateItem(0)

					mailItem.To = "DevExpress"
					mailItem.Subject = "RichEditControl-generated Mail Message"
					mailItem.Body = richEditControl1.Text

					mailItem.Display()
				Else
					MessageBox.Show("Outlook is not available.")
				End If
			Else
				MessageBox.Show("Automation is not available.")
			End If
		End Sub

		Private Function InitializeOutlook() As Boolean
			Dim outlookName As String = "Outlook.Application"

			Try
				outlook = AutomationFactory.GetObject(outlookName)
				Return True
			Catch e1 As Exception
				Try
					outlook = AutomationFactory.CreateObject(outlookName)
					Return True
				Catch e2 As Exception
					Return False
				End Try
			End Try
		End Function
	End Class
End Namespace
