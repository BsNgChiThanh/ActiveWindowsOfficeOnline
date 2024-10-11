# Kích hoạt Windows và Office Bằng Key Online #
- Kích hoạt Windows và Office bằng Key Online.
- ![image](https://github.com/BsNgChiThanh/Lich-phong-kham/assets/82578024/d575f08f-29b1-4848-83b0-fb5e88dcb50c)
- ![image](https://github.com/user-attachments/assets/892ab962-1334-4126-9b74-42be48da0f04)
- Khi kích hoạt Windows bằng key by phone, các bạn đôi khi không getcid được, tôi xin làm bài viết này giúp bạn giải quyết điều đó bằng cách dùng **Windows PowerShell:**
- Đầu tiên bấm tổ hợp phiếm: **Windows + R**, sau đó điền **PowerShell** rồi bấm enter.
- ![image](https://github.com/BsNgChiThanh/Crack-IDM/assets/82578024/73f131a2-efd7-4c50-9a36-106b02d83fca)
- Dán vào câu lệnh: **CD C:\Windows\System32** rồi enter
- ![image](https://github.com/BsNgChiThanh/Crack-IDM/assets/82578024/cc4df65e-6cc1-47a1-a967-fe19d9983a26)

## Nếu active Windows:
- **Dán dòng lệnh dưới đây:**
  
  ```php
  irm https://raw.githubusercontent.com/BsNgChiThanh/ActiveWindowsOfficeOnline/IMP/ActiveWindow_Office_Online.ps1 | iex
  ```

- ![image](https://github.com/user-attachments/assets/009a3f1e-5d28-47b9-b561-195622a7c344)
- Bấm enter, sau đó làm theo chỉ dẫn là xong!

## Nếu active Office:
- **Dán dòng lệnh sau đây:**

  ```php
  irm https://raw.githubusercontent.com/BsNgChiThanh/ActiveWindowsOfficeOnline/IMP/ActiveOfficeBangKeyOnline.ps1 | iex
  ```

- ![image](https://github.com/user-attachments/assets/e34bb824-bce2-4b15-b9fa-856b871c6cd3)
- Bấm enter, sau đó làm theo chỉ dẫn là xong!
  
# NẾU CÓ CHUỔI ID MÀ QUÊN KEY THÌ LÀM NHƯ SAU:

A. ĐỐI VỚI OFFICE:

- Mở **cmd** bằng quyền **Run as Administrator** rồi dán câu lệnh dưới đây vào, nhớ thay **[ID getcid]** trong câu lệnh bên dưới:
 
```PHP
for %a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%a")
if exist "%ProgramFiles% (x86)\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%a"))

cscript.exe OSPP.vbs /actcid:[ID getcid]
cscript.exe OSPP.vbs /act
```

B. ĐỐI VỚI WINDOWS:

- Mở **cmd** bằng quyền **Run as Administrator** rồi dán câu lệnh dưới đây vào, nhớ thay **[ID getcid]** trong câu lệnh bên dưới:
 
```PHP
```

- Done!
