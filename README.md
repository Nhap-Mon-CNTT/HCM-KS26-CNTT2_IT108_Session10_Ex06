# Yêu cầu đầu ra:

## Phần 1:
- Bảng excel xấu là một bảng vị phạm 5 quy tắc vàng:
 
  + Header Duy Nhất:
Hàng đầu tiên là tên các Field. Tuyệt đối không merge ô ở hàng 1.

  + Một Hàng = Một Đối Tượng:
Mỗi dòng chứa thông tin của 1 nhân sự hoặc 1 hóa đơn độc lập.

  + Dữ Liệu Nguyên Tử:
Tách Họ và Tên thành 2 cột riêng để dễ dàng sắp xếp theo ABC.

  + Không Hàng Trống:
Không chèn dòng trống giữa dữ liệu làm đứt quãng vùng nhận diện.

  + Dòng Tổng Nằm Ngoài:
Dòng Total/Average phải đặt riêng biệt, không chen giữa dữ liệu.

  + Mẹo: Hàm =TRIM():
Dùng =TRIM(A2) xóa khoảng trắng thừa vô hình trước khi lọc.
## Phần 2:
Word: [SmartHome_Word.docx](SmartHome_Word.docx)

Excel: [SmartHome_Excel.xlsx](SmartHome_Excel.xlsx)

Powerpoint: [SmartHome_PowerPoint.pptx](SmartHome_PowerPoint.pptx)

## Phần 3:
Nguyên nhân: vùng tra Hoa_hong!A4:B6 là tham chiếu tương đối, nên khi kéo xuống Excel dịch vùng theo từng dòng (A5:B7, A6:B8...). Hạng cần tra rơi khỏi vùng, VLOOKUP trả lỗi #N/A.

Cách sửa: dùng dấu $ để khóa vùng tra.

=VLOOKUP(H5,Hoa_hong!$A$4:$B$6,2,FALSE)

$A$4:$B$6 khóa cả cột và dòng, nên vùng tra đứng yên khi kéo.
H5 không có $, vì giá trị cần tra phải đổi theo từng dòng.
Phím tắt: chọn vùng rồi bấm F4.

Kiểm thử: kéo công thức xuống cuối bảng rồi đối chiếu các dòng giữa và cuối, không chỉ dòng đầu (dòng đầu luôn đúng). Sheet Kiem_thu trong SmartHome_Excel.xlsx cho thấy kết quả: không có $ thì 3 trong 4 dòng lỗi, có $ thì cả 4 dòng đúng.