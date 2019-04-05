<!-- default file list -->
*Files to look at*:

* [MainPage.xaml](./CS/MainPage.xaml) (VB: [MainPage.xaml](./VB/MainPage.xaml))
* [MainPage.xaml.cs](./CS/MainPage.xaml.cs) (VB: [MainPage.xaml.vb](./VB/MainPage.xaml.vb))
<!-- default file list end -->
# DXRichEdit for Silverlight: Out-of-Browser (OOB) application with elevated permissions


<p>This example illustrates how to use RichEditControl in the context of the <a href="http://www.silverlightshow.net/items/Silverlight-4-elevated-permissions.aspx"><u>OOB application with elevated permissions</u></a>. In particular, due to the features available when your application runs inside a relaxed sandbox, you can do the following:</p><p>- Download the content (the image in this example) for the RichEditControl from the remote host regardless of the appropriate <strong>clientaccesspolicy.xml</strong> file presence on this host (see the <a href="https://www.devexpress.com/Support/Center/p/E3484">DXRichEdit for Silverlight: How to force images from external hosts to be loaded into RichEditControl</a> example to learn more on this subject).</p><p>- Load and save the RichEditControl document from/to user folders in the local file system.</p><p>- Email RichEditControl document via the <strong>Outlook.Application</strong> COM object API.</p><p>Also, it should be noted that the "Do you want to allow this application to access your clipboard?" dialog (see <a href="https://www.devexpress.com/Support/Center/p/Q341801">getting clipboard warning dialog while typing in richedit</a>) will not appear in such an application.</p>

<br/>


