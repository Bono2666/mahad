from .models import Checklist


def checklist_notification(request):
    checklists = Checklist.objects.filter(checklist_status='Belum Dikerjakan')

    return len(list(checklists)) if checklists else 0
