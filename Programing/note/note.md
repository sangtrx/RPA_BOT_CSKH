1. Cần sửa lại thời gian lấy trạng thái ở hàm "take_message_status()" trong module "zalo_utility.py"
 + Xét trường hợp 23h59m (ví dụ ngày 26/08) gửi tin và 00h01m (ngày 27/08) lấy trạng thái ==> ngày lấy trạng thái sẽ bị sai lệch
 + Kết quả sai : chat-date = "23:59 - 27/08" ==> cần chỉnh lại "23:59 - 26/08"