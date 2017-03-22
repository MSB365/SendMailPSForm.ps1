 <#
.SYNOPSIS
SendMailPSForm.ps1 

.DESCRIPTION 



.PARAMETER



.EXAMPLE
.\SendMailPSForm.ps1 


.NOTES
You have to edit the default Body Text before you run the Script (Started from Line 95)
Written by: Drago Petrovic

.TROUBLENOTES



Find me on:

* LinkedIn:	https://www.linkedin.com/in/drago-petrovic-86075730/
* Xing:     https://www.xing.com/profile/Drago_Petrovic
* Website:  https://blog.abstergo.ch
* GitHub:   https://github.com/MSB365


Change Log
V1.00, 15/01/2017 - Initial version
V1.01, 17/01/2017 - Fix send permission problems | modify "cancel" Button
V1.10, 23/01/2017 - Change Mail send Method


--- keep it simple, but significant ---

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>


# --- Mail server informations
<# Optional
$SMTPServer = "<EXC-Server>"
$SMTPPort = "25"
$From = "John.doe@contoso.com" 
$Attachment = "C:Path\file.txt"
#>



# --- Script     
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="MyForm" 
    $MyForm.Size = New-Object System.Drawing.Size(750,600) 
     
 
        $mTitle = New-Object System.Windows.Forms.Label 
                $mTitle.Text="New Employee" 
                $mTitle.Top="120" 
                $mTitle.Left="50" 
                $mTitle.Anchor="Left,Top" 
        $mTitle.Size = New-Object System.Drawing.Size(100,50) 
        $MyForm.Controls.Add($mTitle) 
         

# ------------------------------------------------------------------------------------------------------
 
        $mButtonS = New-Object System.Windows.Forms.Button 
                $mButtonS.Text="Send Mail" 
                $mButtonS.Top="520" 
                $mButtonS.Left="520" 
                $mButtonS.Anchor="Left,Top" 
        $mButtonS.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mButtonS) 

        function OnClick($Sender, $EventArgs){
        
 # --- Mail Body Text
$mBody = "Hello $($mName.text),
Here your Company informations:

Username: $($mUNNSurename.text)@iz00.net
Username AD: contoso\$($mUNSurenameN.text)
E-Mail Address: $($mUNNSurename.text)@Contoso.com

Mailserver: https://owa.contoso.com
Office365: https://portal.office.com

Login: $($mUNNSurename.text)@contoso.com
Login Skype for Business: $($mUNNSurename.text)@contoso.com

Your IZ Contact informations to reach you, are:
Phone (SfB): +41 43 123 45 $($mSfBNumber.text)
Mail: $($mUNNSurename.text)@Contoso.com



If you have any questions left, feel free to contact me!


Kind regards

Your Contoso Admin Team


______________________________________________________________________________________
Contoso Switzerland AG | Fakestreet 109 | 8000 Zuerich - Switzerland
Direct: +41 43 123 45 00 | Mobile: +41 43 123 45 67
Phone: +41 43 123 45 67 | Fax: +41 43 123 45 67
noreply@contoso.com | www.contoso.com
______________________________________________________________________________________"       
        
        $Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "$mTO"
$Mail.CC = "$mCC"
$Mail.BCC = "$mBCC"
$Mail.Subject = "$($mMailSubject.text)"
$Mail.Body ="$mBody"
$Mail.Send()
$MyForm.close()
}

        $mButtonS.add_click({OnClick $this $_})

         
# ------------------------------------------------------------------------------------------------------

 
        $mButtonC = New-Object System.Windows.Forms.Button 
                $mButtonC.Text="Cancel" 
                $mButtonC.Top="520" 
                $mButtonC.Left="400" 
                $mButtonC.Anchor="Left,Top" 
        $mButtonC.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mButtonC) 

function OnClick1($Sender, $EventArgs){$myform.close()}

 $mButtonC.add_click({OnClick1 $this $_})


# ------------------------------------------------------------------------------------------------------


        $mButtonS.add_click({OnClick $this $_})
        $mName = New-Object System.Windows.Forms.RichTextBox 
                $mName.Text="Name" 
                $mName.Top="170" 
                $mName.Left="50" 
                $mName.Anchor="Left,Top" 
        $mName.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mName) 
         
 
        $mSurename = New-Object System.Windows.Forms.RichTextBox 
                $mSurename.Text="Surename" 
                $mSurename.Top="200" 
                $mSurename.Left="50" 
                $mSurename.Anchor="Left,Top" 
        $mSurename.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mSurename) 
         
 
        $mUNSurenameN = New-Object System.Windows.Forms.RichTextBox 
                $mUNSurenameN.Text="AD Username (SurenamdN)" 
                $mUNSurenameN.Top="230" 
                $mUNSurenameN.Left="50" 
                $mUNSurenameN.Anchor="Left,Top" 
        $mUNSurenameN.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mUNSurenameN) 
         
 
        $mUNNSurename = New-Object System.Windows.Forms.RichTextBox 
                $mUNNSurename.Text="Username (N.Surename)" 
                $mUNNSurename.Top="260" 
                $mUNNSurename.Left="50" 
                $mUNNSurename.Anchor="Left,Top" 
        $mUNNSurename.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mUNNSurename) 
         
 
        $mSfBNumber = New-Object System.Windows.Forms.RichTextBox 
                $mSfBNumber.Text="SfB Short Number (00)" 
                $mSfBNumber.Top="290" 
                $mSfBNumber.Left="50" 
                $mSfBNumber.Anchor="Left,Top" 
        $mSfBNumber.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mSfBNumber) 
         
 
        $mMailSubject = New-Object System.Windows.Forms.RichTextBox 
                $mMailSubject.Text="Subject" 
                $mMailSubject.Top="170" 
                $mMailSubject.Left="400" 
                $mMailSubject.Anchor="Left,Top" 
        $mMailSubject.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mMailSubject) 
         
 
        $mTO = New-Object System.Windows.Forms.RichTextBox 
                $mTO.Text="Mail address: TO" 
                $mTO.Top="200" 
                $mTO.Left="400" 
                $mTO.Anchor="Left,Top" 
        $mTO.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mTO) 
         
 
        $mCC = New-Object System.Windows.Forms.RichTextBox 
                $mCC.Text="Mail address:CC" 
                $mCC.Top="230" 
                $mCC.Left="400" 
                $mCC.Anchor="Left,Top" 
        $mCC.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mCC) 
         
 
        $mBCC = New-Object System.Windows.Forms.RichTextBox 
                $mBCC.Text="Mail address BCC" 
                $mBCC.Top="260" 
                $mBCC.Left="400" 
                $mBCC.Anchor="Left,Top" 
        $mBCC.Size = New-Object System.Drawing.Size(300,23) 
        $MyForm.Controls.Add($mBCC) 
        $MyForm.ShowDialog()

