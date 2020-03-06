		Set objEmail = CreateObject("CDO.Message")
		objEmail.From = "SSV-Schancheholen-Sørmarka@sweco.no"
		objEmail.To = "samuel.andersson@sweco.no;"
		objEmail.Cc = "thomas.ostgulen@sweco.no"
		objEmail.Subject = "SSV: Oppdaterte fagmodeller"
		objEmail.CreateMHTMLBody "file://C:\Scripts\SSV\Python\htmldoc.html"		
	objEmail.Configuration.Fields.Item _
		    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objEmail.Configuration.Fields.Item _
		    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
		        "printer.swecogroup.com"
		objEmail.Configuration.Fields.Item _
		    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objEmail.Configuration.Fields.Update
		objEmail.Send