"""Script cho template sample_template.docx
Các hàm trong file này sẽ được gọi cho trường type=script.
"""

import random
def tong_gia_tri(don_gia: float, so_luong: int) -> float:
    """Tính tổng giá trị từ đơn giá và số lượng."""
    return don_gia * so_luong * random.randrange(1, 100)
