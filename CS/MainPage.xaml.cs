using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using DevExpress.XtraRichEdit.API.Native;
using System.Runtime.InteropServices.Automation;

namespace RichEditOOBElevatedPermissions {
    public partial class MainPage : UserControl {
        private bool featureComplete = Application.Current.HasElevatedPermissions;
        private const string errorMessage = "This application is not trusted.";
        private dynamic outlook;

        public MainPage() {
            InitializeComponent();
        }

        private void btnLoadImage_Click(object sender, RoutedEventArgs e) {
            if (featureComplete) {
                richEditControl1.Document.InsertImage(richEditControl1.Document.Range.End,
                    DocumentImageSource.FromUri("http://www.devexpress.com/Home/i/logos/preview.png", richEditControl1));
            }
            else {
                MessageBox.Show(errorMessage);
            }
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e) {
            if (featureComplete) {
                string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string fileName = "test.rtf";
                string path = Path.Combine(myDocuments, fileName);

                if (File.Exists(path)) {
                    richEditControl1.RtfText = File.ReadAllText(path);
                }
                else {
                    MessageBox.Show(string.Format("The '{0}' file does not exist.", path));
                }
            }
            else {
                MessageBox.Show(errorMessage);
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e) {
            if (featureComplete) {
                string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string fileName = "test.rtf";
                string path = Path.Combine(myDocuments, fileName);

                File.WriteAllText(path, richEditControl1.RtfText);
            }
            else {
                MessageBox.Show(errorMessage);
            }
        }

        private void btnEmail_Click(object sender, RoutedEventArgs e) {
            if (AutomationFactory.IsAvailable) {
                if (InitializeOutlook()) {
                    dynamic mailItem = outlook.CreateItem(0);

                    mailItem.To = "DevExpress";
                    mailItem.Subject = "RichEditControl-generated Mail Message";
                    mailItem.Body = richEditControl1.Text;

                    mailItem.Display();
                }
                else {
                    MessageBox.Show("Outlook is not available.");
                }
            }
            else {
                MessageBox.Show("Automation is not available.");
            }
        }

        private bool InitializeOutlook() {
            string outlookName = "Outlook.Application";

            try {
                outlook = AutomationFactory.GetObject(outlookName);
                return true;
            }
            catch (Exception) {
                try {
                    outlook = AutomationFactory.CreateObject(outlookName);
                    return true;
                }
                catch (Exception) {
                    return false;
                }
            }
        }
    }
}
