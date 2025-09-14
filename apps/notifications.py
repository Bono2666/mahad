import datetime
from .models import Attendance, Checklist, RegionDetail, Task


def checklist_notification(request):
    all_tasks = Task.objects.all()
    checklist = Checklist.objects.filter(checklist_date=datetime.date.today(
    )) if Checklist.objects.filter(checklist_date=datetime.date.today()) else None
    day = datetime.date.today().weekday()

    if not checklist:
        for task in all_tasks:
            region = RegionDetail.objects.get(
                room_id=task.room_id) if RegionDetail.objects.filter(room_id=task.room_id) else None
            if region:
                new_checklist = Checklist(
                    room_id=task.room_id,
                    janitor=region.region.janitor,
                    task=task,
                    checklist_date=datetime.date.today(),
                )
            else:
                new_checklist = Checklist(
                    room_id=task.room_id,
                    task=task,
                    checklist_date=datetime.date.today(),
                )

            if day == 0 and task.mon:
                new_checklist.save()
            elif day == 1 and task.tue:
                new_checklist.save()
            elif day == 2 and task.wed:
                new_checklist.save()
            elif day == 3 and task.thu:
                new_checklist.save()
            elif day == 4 and task.fri:
                new_checklist.save()
            elif day == 5 and task.sat:
                new_checklist.save()
            elif day == 6 and task.sun:
                new_checklist.save()

    checklists = Checklist.objects.filter(
        checklist_date=datetime.date.today()).exclude(checklist_status='Selesai')

    return len(list(checklists)) if checklists else 0


def urgent_notification(request):
    urgent = Checklist.objects.filter(checklist_urgent=True)

    return len(list(urgent)) if urgent else 0


def attendance(request):
    attendance = Attendance.objects.get(user_id=request.user.user_id, absence_date=datetime.date.today(
    )) if Attendance.objects.filter(user_id=request.user.user_id, absence_date=datetime.date.today()) else None
    return attendance
