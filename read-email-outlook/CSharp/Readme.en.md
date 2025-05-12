[![vi](https://img.shields.io/badge/Ng%C3%B4n%20ng%E1%BB%AF-Ti%E1%BA%BFng%20Vi%E1%BB%87t-red.svg)](https://github.com/ngochoaitn/little-things/blob/main/read-email-outlook/CSharp/Readme.md)
[![en](https://img.shields.io/badge/Language-English-blue.svg)](https://github.com/ngochoaitn/little-things/blob/main/read-email-outlook/CSharp/Readme.en.md)

# C# Code to Read Emails and Get OAuth2 Access Token for Outlook
⚠️**Note**: Some methods have been marked as [Obsolete], such as GetEmailsByRequest and GetOAuth2Token. Use them with caution, as they may trigger account verification and potentially lead to account suspension.
## Example: Get Access Token:
```
string data = "email@outlook.com.vn|password|refresh_token|client_id";
OutlookHelper outlookHelper = new OutlookHelper(data);
string accessToken = outlookHelper.GetAccessToken();
```

## Example: Read Email List:
```
string data = "email@outlook.com.vn|password|refresh_token|client_id";
OutlookHelper outlookHelper = new OutlookHelper(data);
var mails = outlookHelper.GetEmails();
```

## Với new OutlookHelper(data) hỗ trợ sẵn các định dạng:
```
// 1 => email|pass|refreshToken|clientId
// 2 => email|refreshToken|clientId
// 3 => email|pass                       // May trigger account verification
```

![Screenshot](https://github.com/ngochoaitn/little-things/blob/main/read-email-outlook/CSharp/Screenshot.png)