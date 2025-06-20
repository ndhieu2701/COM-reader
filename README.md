COM-PORT reader cho cân điện tử UPA-Q trên window
* đọc dữ liệu cổng com ra dạng text dạng như app 232key:
- dùng python
* cách cài đặt:
- pip install -r requirements.txt

hướng dẫn sử dung:
+ chọn cổng COM, nếu không có thử bấm nút refresh cổng COM để xem có không
+ điền lựa chọn Baud, Encoding
+ dùng regex để lọc lấy kí tự mình cần
+ có lựa chọn ấn Enter sau khi lấy được không, tự động khởi động cùng window không
+ timeout để check xem trong khoảng thời gian đó nếu không nhận được dữ liệu sẽ tự động giải phóng cổng COM và reconnect lại sau 5s :v
+ ghi log và clear log

có 2 version
- v2 là giao diện đẹp hơn

custom từ repo: https://github.com/Smartlux/py232key
