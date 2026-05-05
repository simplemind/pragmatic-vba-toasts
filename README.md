<div>
  <h1>Pragmatic VBA toast notifications</h1>
  <p>I could not find working code so decided to build it. Perhaps it will come handy to others too. Enjoy.</p>
  <p>The code tries to avoid fragile complexities that often break in VBA. The code is designed to work on Windows.</p>
  <p>
    <h2>To use</h2>
  <ul>
    <li>Download the files. Click on Code -> Download ZIP</li>  
    <li>Unzip the files</li>
    <li>Unblock the default block on downloaded VBA files: Right click on the Excel file and files with extensions frm and bas -> Show more options -> Properties -> Tick near Unblock.</li>
  </ul>
</p>
<p>A demo can be seen in the enclosed Excel file. Click once or several times on the green Test Toat button to see toast notifications.</p>
<p>To setup the toast notifications in your workbook you need to copy the form and modules. Then  call notifications using ShowToast</p>
<h2>How to copy the modules</h2>
<div>
  <p>There are two ways to copy the modules:</p>
  <ol>
    <li>Open your Excle workbook, go to VBA Editor (Alt+F11 or Developer -> Visual Basic)</li>
    <h3>1. Drag & drop</h3>
    <li>Within the VBA Editor drag the files frmToast.frm, modMain.bas, modToastService.bas, modWindowEffects.bas on the name of your workbook. I.e. within the left menu.</li>
  <h3>2.Import</h3>
    <li>Alternatively right click on the name of your workbook within the VVBA Editor,</li>
    <li>click on Import File,</li>
    <li>locate one at a time files frmToast.frm, modMain.bas, modToastService.bas, modWindowEffects.bas and click Open.</li>
    <p>This will import just like the drag and drop method. Now you can use the toast notifications by calling ShowToast procedure. See below for details.</p> 
  </ol>
</div>

<h3>Notifications are called using </h3>
<div>
 <div> 
  <code>ShowToast notification-title, notification-message, [optional duration]</code>
  <br />or<br />
  <code>Call ShowToast(notification-title, notification-message, [optional duration]</code>)
  <br />
 </div>
  <br />
  For example:
   <div>
  <code>ShowToast, "Confirmation", "File has been successfully uploaded"</code>
  <br />
  <code>ShowToast, "Confirmation", "File has been successfully uploaded", 4</code>
  </div>
  <br />or using a Call method<br />
  <code>Call ShowToast("Confirmation", "File has been successfully uploaded")</code>
  <br />
  <code>Call ShowToast("Confirmation", "File has been successfully uploaded")</code>
  </div>
</div>
<p>
  <h2>Key</h2>
  <ul>
    <li><b>notification-title</b> - a title that will be displayed at the top of your notification</li>
    <li><b>notification-message</b> - notification message displayed below the title</li>
    <li><b>duration</b> - an optional parameter. A number of seconds to show the notification before it is closed. Must use whole numbers. Defaults to 3 seconds, if not provided. </li>
  </ul>
</p>

</div>
