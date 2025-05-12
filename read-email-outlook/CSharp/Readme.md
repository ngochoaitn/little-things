[![vi](https://img.shields.io/badge/Ng%C3%B4n%20ng%E1%BB%AF-Ti%E1%BA%BFng%20Vi%E1%BB%87t-red.svg)](https://github.com/ngochoaitn/little-things/blob/main/read-email-outlook/CSharp/Readme.md)
[![en](https://img.shields.io/badge/Language-English-blue.svg)](https://github.com/ngochoaitn/little-things/blob/main/read-email-outlook/CSharp/Readme.en.md)

# Code C# đọc email, lấy access token OAuth2 Outlook
⚠️**Chú ý**: Có một số hàm đã được đánh dấu [Obsolete] như: GetEmailsByRequest, GetOAuth2Token, cần cẩn trọng khi sử dụng vì có thể sẽ bị xác minh tài khoản dẫn tới khóa tài khoản 
## Ví dụ lấy access token:
```
string data = "email@outlook.com.vn|password|refresh_token|client_id";
OutlookHelper outlookHelper = new OutlookHelper(data);
string accessToken = outlookHelper.GetAccessToken();
```

## Ví dụ đọc danh sách email:
```
string data = "email@outlook.com.vn|password|refresh_token|client_id";
OutlookHelper outlookHelper = new OutlookHelper(data);
var mails = outlookHelper.GetEmails();
```

## Với new OutlookHelper(data) hỗ trợ sẵn các định dạng:
```
// 1 => email|pass|refreshToken|clientId
// 2 => email|refreshToken|clientId
// 3 => email|pass                       // Có thể bị xác minh tài khoản khoản
```