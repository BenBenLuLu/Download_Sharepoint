# Download_Sharepoint
IT會檔Sharepoint 的權限， 目前沒有辦法
你現在需要做的事
1. 請 IT 幫忙註冊一個 Azure AD App
專案裡已新增 AZURE_APP_SETUP.md，可以直接轉給 IT。

IT 需要設定：

項目	設定
Platform	Mobile and desktop → http://localhost
Public client flows	Yes
API Permissions	Files.Read.All + Sites.Read.All（Delegated）
Admin consent	必須同意
2. IT 給你 Client ID 後
貼到程式 Step 2 → Client ID 欄位
Tenant 維持 shlgroup.onmicrosoft.com
點 Sign in with Microsoft
瀏覽器登入成功後 → Start Download
