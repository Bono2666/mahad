from .models import Checklist


def checklist_notification(request):
    checklists = Checklist.objects.all().exclude(checklist_status='Selesai')

    return len(list(checklists)) if checklists else 0
