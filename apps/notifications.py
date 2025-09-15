import datetime
from .models import Attendance


def attendance(request):
    attendance = Attendance.objects.get(user_id=request.user.user_id, absence_date=datetime.date.today(
    )) if Attendance.objects.filter(user_id=request.user.user_id, absence_date=datetime.date.today()) else None
    return attendance
