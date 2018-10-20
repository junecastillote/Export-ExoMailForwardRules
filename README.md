<p>
Main Post: https://www.lazyexchangeadmin.com/2018/10/office-365-mailbox-forwarding-rules.html <br />
Being on top of who’s forwarding messages to who’s email, especially those being forwarded to external domains is essential to email security for administrators. Phishing attacks can leave your users’ mailboxes prone to data exfiltration by way of forwarding emails, and so being able to regularly review and audit mailbox forwarding rules is beneficial to protecting your company’s data.</p>
<p>
This script can be used to export a report of all the forward/redirect rules present in all user mailboxes.</p>
<h3>
Download Link</h3>
<p>
<a title="https://github.com/junecastillote/Export-ExoMailForwardRules" href="https://github.com/junecastillote/Export-ExoMailForwardRules">https://github.com/junecastillote/Export-ExoMailForwardRules</a></p>
<h3>
Requirements</h3>
<ul>
<li>Must have an Office 365 account that is assigned at least an Exchange Administrator role whose credentials will be used to connect to Office 365 PowerShell.</li>
<ul>
<li>It is important that the account is not MFA enabled as the script operates by paging and re-authenticates to Office 365 page.</li>
</ul>
<li>Must have a mailbox to be able to send the email report using Office 365 SMTP Relay. This could be the Service Account you’re using for the session, or a Shared Mailbox that the Service Account has Send As permission to. If you do not plan to send the report thru email, then you can disregard this requirement.</li>
</ul>
<h3>
How to use</h3>
<h4>
Setup Office 365 Credentials</h4>
<ul>
<li>Open PowerShell and change to the directory where the script is saved (eg. C:\Scripts\Export-ExoMailForwardRules)</li>
<li>Run this command:</li>
<p>
<strong>Get-Credential | Export-CliXml Office365StoredCredential.xml</strong><p>
<strong><a href="https://lh3.googleusercontent.com/-MropFbRbo78/W8s_hiNYNII/AAAAAAAADWE/qrb5GDahihksn1yGer_BKaKRJAkfpyGjACHMYCw/s1600-h/mRemoteNG_2018-10-20_21-33-11%255B2%255D"><img width="733" height="337" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-TIsTTJg70JU/W8s_jG8XeAI/AAAAAAAADWI/T2IiMnsAQSQv_uQvyarhd7NpgJYEmhJVgCHMYCw/mRemoteNG_2018-10-20_21-33-11_thumb?imgmax=800" border="0"></a></strong>
<li>This saves the encrypted credential in the same folder</li>
</ul>
<h4>

</h4>
<h4>
Modify Variables</h4>
<h4>
Email Settings</h4>
<p>
<a href="https://lh3.googleusercontent.com/-JieQOVwidUc/W8s_jytOGhI/AAAAAAAADWM/VEhc8sm5dBo-TzLm_M52AKAniIaonO7LQCHMYCw/s1600-h/mRemoteNG_2018-10-20_22-33-35%255B2%255D"><img width="542" height="140" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-4EjuGi8CZLg/W8s_lHLQCRI/AAAAAAAADWQ/1QNQ2LzROccT8NxyRsUD2qILlU2vHsi2wCHMYCw/mRemoteNG_2018-10-20_22-33-35_thumb?imgmax=800" border="0"></a></p>
<p>
<strong>NOTE</strong>: The <strong><em>$sender</em></strong> value must be the actual email address of the service account or the shared mailbox used for sending the email report.</p>
<h4>
Paging</h4>
<p>
In cases where there are a large number of mailboxes to be processed, the Exchange Online PowerShell session may timeout/disconnect which would cause the script to fail. As a workaround, this script is configured to process the mailboxes in pages. By default, the page settings is set to 100 – which means after every 100 mailboxes processed, the script will re-establish and re-authenticate the PowerShell session. You can increase the page value but it is not recommended to set it too high.</p>
<p>
<a href="https://lh3.googleusercontent.com/--uCA58YUJso/W8tCV7NM8oI/AAAAAAAADXI/XtcgfjQcWqwaUWmfvbc5U0QCoMQ7ib6jACHMYCw/s1600-h/mRemoteNG_2018-10-20_22-51-45%255B2%255D"><img width="122" height="42" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-KaQIYtUSA2o/W8tCXi1v4rI/AAAAAAAADXM/IkYwP0OaPcwYoSvx79nDx1EizP2HnURyACHMYCw/mRemoteNG_2018-10-20_22-51-45_thumb?imgmax=800" border="0"></a>&nbsp;</p>
<h4>
Run the script</h4>
<p>
The script requires no parameters.</p>
<p>
<a href="https://lh3.googleusercontent.com/-1F3qRx4RNgU/W8s_l1ajNSI/AAAAAAAADWU/aMtiHEZWK2cxer2-tUX00XVTc4UPecK_wCHMYCw/s1600-h/mRemoteNG_2018-10-20_21-54-17%255B7%255D"><img width="742" height="300" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-DJRfyRUVXQg/W8s_nLi1LnI/AAAAAAAADWY/BqCdggEuqqYkDmtzLcC1vncZ4ZplqioqwCHMYCw/mRemoteNG_2018-10-20_21-54-17_thumb%255B3%255D?imgmax=800" border="0"></a></p>
<h3>
Output</h3>
<p>
CSV File</p>
<p>
The csv file gets saved in the “\Reports” folder</p>
<p>
<a href="https://lh3.googleusercontent.com/-iFfFniyJFJA/W8s_oKDeprI/AAAAAAAADWc/mRi3jRQCTaYfCr2j7kn2YOSAb17rhzttwCHMYCw/s1600-h/mRemoteNG_2018-10-20_22-43-04%255B2%255D"><img width="457" height="90" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-DXBqakNxA9U/W8s_pMsyupI/AAAAAAAADWg/hTQtFjRJ3GIVqGj3yaT-iWkLv3ZQlV6qQCHMYCw/mRemoteNG_2018-10-20_22-43-04_thumb?imgmax=800" border="0"></a></p>
<p>
<a href="https://lh3.googleusercontent.com/-kdEbNWgtak8/W8tCYSDc5tI/AAAAAAAADXQ/uhAGJ34qRfcOXJVmjFlJ634vVmmmOz4TACHMYCw/s1600-h/mRemoteNG_2018-10-20_22-55-56%255B2%255D"><img width="820" height="70" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-fiIRGipNvmE/W8tCZ6Ye2eI/AAAAAAAADXU/inR03U7Mk_U269ZB81HubbzAHJ2myPKZgCHMYCw/mRemoteNG_2018-10-20_22-55-56_thumb?imgmax=800" border="0"></a></p>
<h4>
Email</h4>
<p>
<a href="https://lh3.googleusercontent.com/-Io8Zgwezkfo/W8s_qKf5OMI/AAAAAAAADWk/tqcFZyXAOzMCFL0_WSWqtHisWvfuy8xVwCHMYCw/s1600-h/mRemoteNG_2018-10-20_22-44-18%255B2%255D"><img width="470" height="207" title="" style="display: inline; background-image: none;" alt="" src="https://lh3.googleusercontent.com/-fcoKymFz-dE/W8s_rFr1hBI/AAAAAAAADWo/y_xXAupJlBoduxh7PJoQZe6CfoJ8tMdygCHMYCw/mRemoteNG_2018-10-20_22-44-18_thumb?imgmax=800" border="0"></a></p>
