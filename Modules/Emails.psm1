function Send-disabledStaff ($To, $From, $SMTPServer, $DisabledUser)
{
    $DisabledUser=($disabledUser -join "<br>")

    $EmailBody = @"
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        body {
                background-color: #f6f6f6;
                font-family: sans-serif;
                -webkit-font-smoothing: antialiased;
                font-size: 14px;
                line-height: 1.4;
                margin: 0;
                padding: 0; 
                -ms-text-size-adjust: 100%;
                -webkit-text-size-adjust: 100%; }

         .wrapper {
                box-sizing: border-box;
                padding: 20px; 
        }
        .main {
                background: #fff;
                border-radius: 3px;
                width: 100%; }
      
              /* -------------------------------------
                  TYPOGRAPHY
              ------------------------------------- */
              h1,
              h2 {
        font-size: 20px;
             color: #247454;
                font-family: sans-serif;
                font-weight: 500;
                line-height: 1;
                margin: 0;
                Margin-bottom: 30px;
          }
           }
              h3,
              h4 {
                color: #000000;
                font-family: sans-serif;
                font-weight: 500;
                line-height: 1.4;
                margin: 0;
                Margin-bottom: 30px; }
              h1 {
                font-size: 25px;
                font-weight: 300;
                text-align: center;
                text-transform: capitalize; }
              p,
              ul,
              ol {
                font-family: sans-serif;
                font-size: 14px;
                font-weight: normal;
                margin: 0;
                Margin-bottom: 15px; }
                p li,
                ul li,
                ol li {
                  list-style-position: inside;
                  margin-left: 5px; }
              a {
                color: #3498db;
                text-decoration: underline; }
    
        </style>
        </head>
        <body>
        <table class="main">
                      <tr>
                        <td class="wrapper">
                          <table border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td>
                                <p><h1>Inactive Staff Accounts have been Disabled.</h1></p>
                        
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                                <p>
        <p><h2>Disabled Users(s):</h2> <ol>$DisabledUsers</ol></p>

                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
              
        </body>
        </html> 
"@
     
        send-MailMessage -to "$To" -from "$From" -subject "Disabled Staff Accounts"  -SmtpServer "$SMTPServer" -Body "$EmailBody" -bodyashtml -verbose
}

Function Write-StudentWelcomeLetter
{
 Param 
    (
        [Parameter(Mandatory=$true)][string]$DisplayName, 
        [Parameter(Mandatory=$true)][string]$CASESID, 
        [Parameter(Mandatory=$true)][string]$UPN,
        [Parameter(Mandatory=$true)][string]$userPassword,
        [Parameter(Mandatory=$true)][string]$TempDirectory
    )

$WelcomeLetter=@"
<!--Header-->
<!DOCTYPE html>
<html>

<head>
   <style>
        .heading {
            top: 0;
            left: 0;
            width: 100%;
            text-align: center;
        }

        .footer {
            color: #999999;
            padding-top: 5px;
            text-align: center;
            width: 100%;
        }

        .comment {
            color: #999999;
            width: 100%;
        }
                .inlineTable {
            display: inline-block;
        }

        BODY {
            font-family: sans-serif;
            font-size: 14px;
            line-height: 1.4;
            color: #4B4C4B;
        }

        TABLE {
            font-size: 12px;
            margin: 5px;
            border: #ccc 1px solid;


        }

        TH {
            padding: 5px;
            background: #00833d;
            color: #ffffff;
            width: 150px;
        }

        TD {
            padding: 5px;
            border-bottom: 1px solid #e0e0e0;
            border-left: 1px solid #e0e0e0;
            width: 150px;
        }

        h1 {
            font-size: 20px;
            color: #00833d;
        }

        h2 {
            font-size: 14px;
            color: #4B4C4B;
            font-weight: 10px;
        }

        h3 {
            font-size: 14px;
            color: #00833d;
        }
    </style>
</head>
<!--Body-->

<body>
    <div class="heading">
        <p>
            <h1>Welcome to ICT at Western Port Secondary College </h1>
        </p>
    </div>

    <body>
        <div class="body">
            <div class="Body">
            <br>
                <h3>Hi $DisplayName,</h3>

                <p>Two accounts have been created for you to access Western Port Secondary College's network and online services. These accounts will provide access to services such as Wi-Fi, Compass, Printing and Office 365.
                    <br>
                    <br>
            </div>
            <table class="inlineTable">
                <tr>
                    <th>Device Username:</th>
                    <td>$CASESID
                        <br>
                    </td>
                </tr>
                <tr>
                    <th>Online Username:</th>
                    <td>$UPN
                        <br>
                    </td>
                </tr>
                <tr>
                    <th>School Password:</th>
                    <td>$userPassword
                        <br>
                    </td>
                </tr>
            </table>

        </div>
        <div class="comment">
        <br>
            Please note this password is case sensitive and must be changed upon logon. To make this change the logon must be to a school managed device such as one of the school laptops and desktops<br>
        </div>
        </div>
        

        <div class="Body">
            <br>
            <h3>Online Services at Western Port Secondary College</h3>
            <h2>Go</h2>
            <p>This website serves as a great starting place to navigate the online services offered at Western Port Secondary College.
                <br><br>
                You can access go from <a href=http://go.westernportsc.vic.edu.au>http://go.westernportsc.vic.edu.au</a>
            </p>
            <br>
            <p>
                    <h2>Compass</h2>
                    Compass is our school administration system.  You can see your class schedule and stay up-to-date with changes to your timetable.<br> <br>
                    You can access Compass from <a href=http://go.westernportsc.vic.edu.au>Go</a> or directly on <a href=https://westernportsc-vic.compass.education/>https://westernportsc-vic.compass.education</a>
            </p>
            <br>
            <h2>Office 365</h2>
            <p>Office 365 is a portal of collaboration software used at Western Port Secondary College. <br>
            Using Office 365 you will have access to Microsoft Word, Excel, Powerpoint, Emails and much more. <br>
            <br>You can access Office 365 from <a href=http://go.westernportsc.vic.edu.au>Go</a> or directly on <a href=http://office.com/>http://office.com</a></p>
            <div class="comment">Accessing Office 365 requires that your eduPass account has been correctly set up by following the eduPass Registration Letter. </div>
        </div>
        <div class="footer">
                <h2>Remember to never share your password. <br>If your password is known to others it can put your account at risk.  <br>Should you think your password is compromised, contact ICT support <u>immediately</u></h2>
                </p>
            </div>
    </body>
    <!--footer-->
    <div class="footer">
        <p>Generated: $(Get-Date -Format dd/MM/yyyy) </p>
    </div>
<br>

     </html>
"@


    $WelcomeLetter > $tempDirectory\$CASESID.html

    LogWrite -logString "Generated Student Welcome Letter"

    return "$tempDirectory\$CASESID.html"

}


function Send-StudentWelcome
{

 Param 
    (
        [Parameter(Mandatory=$true)][string]$DisplayName, 
        [Parameter(Mandatory=$true)][string]$CASESID, 
        [Parameter(Mandatory=$true)][string]$UPN,
        [Parameter(Mandatory=$true)][string]$HomeGroup
    )

$Head =@"
<!DOCTYPE html>
<html>
<head>
    <style>
        .heading {top: 0;left: 0;width: 100%;}
        .footer {color: #999999;padding-top: 5px;text-align: center;width: 100%;}
        BODY {font-family: sans-serif;font-size: 14px;line-height: 1.4;color: #4B4C4B;}
        TABLE {font-size: 12px;margin: 5px;border: #ccc 1px solid;width: 0%;}
        TH {padding: 5px;background: #2980b9;color: #ffffff}
        TD {padding: 5px;border-bottom: 1px solid #e0e0e0;border-left: 1px solid #e0e0e0;}

        h1 {font-size: 20px;color: #2980b9;}

    </style>
</head>
"@
$Body = @"
<body>
    <div class="body">
        <div class="heading">
            <p>
                <h1>New Student Account</h1>
            </p>
        </div>
        <table>
            <tr>
                <th>Student:</th>
                <td>$DisplayName</td>
            </tr>
            <tr>
                <th>Local Login (CASES ID):</th>
                <td>$CASESID</td>
            </tr>
                <th>Online Services Login:</th>
                <td>$UPN</td>
            </tr>
            <tr>
                <th>Homegroup:</th>
                <td>$HomeGroup</td>
            </tr>
        </table>
        <br>
		A 'Welcome Letter' has been attached to this email.<br>
<br>
<strong><u>This letter can be distributed only once a signed ICT Usage Agreement has been returned</u> </strong>.<br>
<br>
    </div>
    </html>
"@
$Foot = @"
<div class="footer">
    <p>Generated: $(Get-Date -Format dd/MM/yyyy) </p>
    <p>By $env:COMPUTERNAME</p>
</div>
"@

return "$Head $Body $Foot"
}

function Send-StudentWelcomeNoAUP
{

 Param 
    (
        [Parameter(Mandatory=$true)][string]$DisplayName, 
        [Parameter(Mandatory=$true)][string]$CASESID, 
        [Parameter(Mandatory=$true)][string]$HomeGroup
    )

    $policyLaptop = ""

    if ( $HomeGroup.Substring(0,2) -eq "07" -or $HomeGroup.Substring(0,2) -eq "08" -or $HomeGroup.Substring(0,2) -eq "10")
    {
        $policyLaptop = "Additionaly as this student is in a year level with a school managed laptop program, if the student has not returned the laptop program acceptance they will not be issued a device. <br>"
    }
    

$Head =@"
<!DOCTYPE html>
<html>
<head>
    <style>
        .heading {top: 0;left: 0;width: 100%;}
        .footer {color: #999999;padding-top: 5px;text-align: center;width: 100%;}
        BODY {font-family: sans-serif;font-size: 14px;line-height: 1.4;color: #4B4C4B;}
        TABLE {font-size: 12px;margin: 5px;border: #ccc 1px solid;width: 0%;}
        TH {padding: 5px;background: #2980b9;color: #ffffff}
        TD {padding: 5px;border-bottom: 1px solid #e0e0e0;border-left: 1px solid #e0e0e0;}

        h1 {font-size: 20px;color: #2980b9;}

    </style>
</head>
"@
$Body = @"
<body>
    <div class="body">
        <div class="heading">
            <p>
                <h1>New Student Account - No AUP Registered</h1>
            </p>
        </div>
        $DisplayName in $HomeGroup has been registered as a new student with the ICT Team.<br>
<br>
However there is no Acceptable use Policy on file, no accounts or services will be available to the student until the AUP has returned. <br>
$policyLaptop
<br>
    </div>
    </html>
"@
$Foot = @"
<div class="footer">
    <p>Generated: $(Get-Date -Format dd/MM/yyyy) </p>
    <p>By $env:COMPUTERNAME</p>
</div>
"@

return "$Head $Body $Foot"
}



function Send-StudentExit
{

 Param 
    (
        [Parameter(Mandatory=$true)][string]$DisplayName, 
        [Parameter(Mandatory=$true)][string]$CASESID, 
        [Parameter(Mandatory=$true)][string]$Days,
        [string]$schoolsEmail
    )


$Head =@"
<!DOCTYPE html>
<html>
<head>
    <style>
        .heading {top: 0;left: 0;width: 100%;}
        .footer {color: #999999;padding-top: 5px;text-align: center;width: 100%;}
        BODY {font-family: sans-serif;font-size: 14px;line-height: 1.4;color: #4B4C4B;}
        TABLE {font-size: 12px;margin: 5px;border: #ccc 1px solid;width: 0%;}
        TH {padding: 5px;background: #2980b9;color: #ffffff}
        TD {padding: 5px;border-bottom: 1px solid #e0e0e0;border-left: 1px solid #e0e0e0;}

        h1 {font-size: 20px;color: #2980b9;}

    </style>
</head>
"@
$Body = @"
<body>
    <div class="body">
        <div class="heading">
            <p>
                <h1>Western Port Secondary College Student Exit </h1>
            </p>
        </div>
        
        Good Day $DisplayName,
        <br>
        The ICT team has been notified that you have left Western Port Secondary College and as such your student account is in the process of being exited. <br>
        Currently you have $Days left to complete the following before your account is removed. <br>

        Upon the removal of your account, anything using services that are authenticated by $CASESID@westernportsc.vic.edu or $CASESID will cease to work and all data will be removed <br>
        in the mean time your account can no longer log onto any school owned device for the remainder of this exit period but your online accounts will still work to allow you to back up any data that you would <br>
        like to, permissions to interact with students and staff have also been removed.<br>
        <br>
        The two most common items to want to backup are your Email and OneDrive and we encourage you to do this as soon as possible as once the account expires the information is simply gone and cannot be retrieved<br>
        <br>
        Additionally anything you have authorised school managed services on through the $CASESID@westernportsc.vic.edu.au account on such as Microsoft Office will also cease to function
        <br>
        
        $extraInfo

        Regards 
        
        Western Port Secondary College ICT Team
<br>
    </div>
    </html>
"@
$Foot = @"
<div class="footer">
    <p>Generated: $(Get-Date -Format dd/MM/yyyy) </p>
    <p>By $env:COMPUTERNAME</p>
</div>
"@

return "$Head $Body $Foot"
}