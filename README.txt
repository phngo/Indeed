Instruction how to install Windows Service
1. Open Visual Studio 2015 and compile project with RELEASE precompiler directive. In case there you don't have Visual Studio, I have included RELEASE build with the repository.
2. Open Command Line as Administrator
3. Navigate to \IndeedSample\IndeedSample\bin\Release
4. Type command "C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe" "IndeedSample.exe"
5. Start SMTPService in services panel 
	a. Ctrl-Alt-Del 
	b. Services 
	c. Locate SMTPService and start it (Right click and Start)
6. Parameters are saved in \IndeedSample\IndeedSample\App.config
	Please change it accordingly
	<appSettings>
		<add key="SMTPserver" value="<required smtp server>"/>
		<add key="SMTPport" value="<required port>"/>
		<add key="SMTPLogin" value="<required login>"/>
		<add key="SMTPpassword" value="<required password>"/>
		<add key="EmailAddress" value="<required To Address>"/>
	</appSettings>