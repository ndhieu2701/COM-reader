COM-PORT reader cho cân điện tử UPA-Q trên window
* đọc dữ liệu cổng com ra dạng text dạng như app 232key:
- dùng python
* cách cài đặt:
- pip install -r requirements.txt
- build ra .exe:
  pyinstaller --onefile --windowed upa_q_v2.py (trong đó upa_q_v2.py là tên file của bạn)
  sau khi build file sẽ nằm trong thư mục dist/
 * build v3: pyinstaller --onefile --windowed upa_q_v3.py --add-data "rs232UsbIcon.png;." --add-data "rs232UsbIcon.ico;."
hướng dẫn sử dung:
+ chọn cổng COM, nếu không có thử bấm nút refresh cổng COM để xem có không
+ điền lựa chọn Baud, Encoding
+ dùng regex để lọc lấy kí tự mình cần
+ có lựa chọn ấn Enter sau khi lấy được không, tự động khởi động cùng window không
+ timeout để check xem trong khoảng thời gian đó nếu không nhận được dữ liệu sẽ tự động giải phóng cổng COM và reconnect lại sau 5s :v
+ ghi log và clear log

có 3 version
- v2 là giao diện đẹp hơn
- v3 giống v2, thêm chức năng chạy ngầm khi tắt cửa sổ

custom từ repo: https://github.com/Smartlux/py232key
