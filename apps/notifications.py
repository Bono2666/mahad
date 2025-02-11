import datetime
from .models import Checklist, RegionDetail, Task


def checklist_notification(request):
    all_tasks = Task.objects.all()
    checklist = Checklist.objects.filter(checklist_date=datetime.date.today(
    )) if Checklist.objects.filter(checklist_date=datetime.date.today()) else None

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
            new_checklist.save()

    checklists = Checklist.objects.filter(
        checklist_date=datetime.date.today()).exclude(checklist_status='Selesai')

    return len(list(checklists)) if checklists else 0


def urgent_notification(request):
    urgent = Checklist.objects.filter(checklist_urgent=True)

    return len(list(urgent)) if urgent else 0
