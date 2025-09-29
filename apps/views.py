import base64
import os
from django.contrib.auth import update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.db import connection, IntegrityError
from django.http import HttpResponseRedirect
from django.shortcuts import render
from django.urls import reverse
from apps.forms import *
from apps.models import *
from authentication.decorators import role_required
from django.http import HttpResponse
import xlsxwriter
from django.db.models import Value, CharField
from django.db.models.functions import Concat, Cast
from . import host
from django.conf import settings
from apps.notifications import *
from geopy.geocoders import Nominatim


@login_required(login_url='/login/')
def home(request):
    context = {
        'attendance': attendance(request),
        'segment': 'index',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
    }
    return render(request, 'home/index.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def user_index(request):
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT user_id, username, email, position_name FROM apps_user INNER JOIN apps_position ON apps_user.position_id = apps_position.position_id")
        users = cursor.fetchall()

    context = {
        'data': users,
        'attendance': attendance(request),
        'segment': 'user',
        'group_segment': 'master',
        'crud': 'index',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
    }

    return render(request, 'home/user_index.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def user_add(request):
    position = Position.objects.all()
    if request.POST:
        form = FormUser(request.POST, request.FILES)
        if form.is_valid():
            form.save()

            if form.instance.signature:
                user = User.objects.get(user_id=form.instance.user_id)
                my_file = user.signature
                filename = '../../www/mahad/apps/media/' + my_file.name
                with open(filename, 'wb+') as temp_file:
                    for chunk in my_file.chunks():
                        temp_file.write(chunk)

            return HttpResponseRedirect(reverse('user-view', args=[form.instance.user_id, ]))
        else:
            message = form.errors
            context = {
                'form': form,
                'position': position,
                'attendance': attendance(request),
                'segment': 'user',
                'group_segment': 'master',
                'crud': 'add',
                'message': message,
                'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
                'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
            }
            return render(request, 'home/user_add.html', context)
    else:
        form = FormUser()
        context = {
            'form': form,
            'position': position,
            'attendance': attendance(request),
            'segment': 'user',
            'group_segment': 'master',
            'crud': 'add',
            'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
            'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
        }
        return render(request, 'home/user_add.html', context)


# View User
@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def user_view(request, _id):
    users = User.objects.get(user_id=_id)
    auth = Auth.objects.filter(user_id=_id)
    form = FormUserView(instance=users)
    position = Position.objects.all()
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT apps_menu.menu_id, menu_name, q_auth.menu_id FROM apps_menu LEFT JOIN (SELECT * FROM apps_auth WHERE user_id = '" + str(_id) + "') AS q_auth ON apps_menu.menu_id = q_auth.menu_id WHERE q_auth.menu_id IS NULL")
        menu = cursor.fetchall()

    if request.POST:
        check = request.POST.getlist('checks[]')
        for i in menu:
            if str(i[0]) in check:
                try:
                    auth = Auth(user_id=_id, menu_id=i[0])
                    auth.save()
                except IntegrityError:
                    continue
            else:
                Auth.objects.filter(user_id=_id, menu_id=i[0]).delete()

        return HttpResponseRedirect(reverse('user-view', args=[_id, ]))

    context = {
        'form': form,
        'formAuth': form,
        'data': users,
        'auth': auth,
        'menu': menu,
        'positions': position,
        'attendance': attendance(request),
        'segment': 'user',
        'group_segment': 'master',
        'tab': 'auth',
        'crud': 'view',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/user_view.html', context)


# Update Auth
@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def auth_update(request, _id, _menu):
    auth = Auth.objects.get(user=_id, menu=_menu)

    if request.POST:
        auth.add = 1 if request.POST.get('add') else 0
        auth.edit = 1 if request.POST.get('edit') else 0
        auth.delete = 1 if request.POST.get('delete') else 0
        auth.save()

        return HttpResponseRedirect(reverse('user-view', args=[_id, ]))

    return render(request, 'home/user_view.html')


# Delete Auth
@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def auth_delete(request, _id, _menu):
    auth = Auth.objects.filter(user=_id, menu=_menu)

    auth.delete()
    return HttpResponseRedirect(reverse('user-view', args=[_id, ]))


@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def remove_signature(request, _id):
    users = User.objects.get(user_id=_id)
    users.signature = None
    users.save()
    return HttpResponseRedirect(reverse('user-view', args=[_id, ]))


# Update User
@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def user_update(request, _id):
    users = User.objects.get(user_id=_id)
    position = Position.objects.all()
    auth = Auth.objects.filter(user_id=_id)

    if request.POST:
        form = FormUserUpdate(request.POST, request.FILES, instance=users)
        if form.is_valid():
            form.save()
            if users.signature:
                my_file = users.signature
                filename = '../../www/mahad/apps/media/' + my_file.name
                with open(filename, 'wb+') as temp_file:
                    for chunk in my_file.chunks():
                        temp_file.write(chunk)

            return HttpResponseRedirect(reverse('user-view', args=[_id, ]))
    else:
        form = FormUserUpdate(instance=users)

    message = form.errors
    context = {
        'form': form,
        'data': users,
        'positions': position,
        'auth': auth,
        'attendance': attendance(request),
        'segment': 'user',
        'group_segment': 'master',
        'crud': 'update',
        'tab': 'auth',
        'message': message,
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/user_view.html', context)


# Delete User
@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def user_delete(request, _id):
    users = User.objects.get(user_id=_id)

    users.delete()
    return HttpResponseRedirect(reverse('user-index'))


@login_required(login_url='/login/')
def change_password(request):
    if request.POST:
        form = FormChangePassword(data=request.POST, user=request.user)
        if form.is_valid():
            form.save()
            update_session_auth_hash(request, form.user)
            return HttpResponseRedirect(reverse('home'))
    else:
        form = FormChangePassword(user=request.user)

    message = form.errors
    context = {
        'form': form,
        'data': request.user,
        'crud': 'update',
        'message': message,
        'attendance': attendance(request),
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
    }
    return render(request, 'home/user_change_password.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='USER')
def set_password(request, _id):
    users = User.objects.get(user_id=_id)
    if request.POST:
        form = FormSetPassword(data=request.POST, user=users)
        if form.is_valid():
            form.save()
            update_session_auth_hash(request, form.user)
            return HttpResponseRedirect(reverse('user-view', args=[_id, ]))
    else:
        form = FormSetPassword(user=users)

    message = form.errors
    context = {
        'form': form,
        'data': users,
        'attendance': attendance(request),
        'segment': 'user',
        'group_segment': 'master',
        'crud': 'update',
        'message': message,
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='USER') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/user_set_password.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='POSITION')
def position_add(request):
    if request.POST:
        form = FormPosition(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return HttpResponseRedirect(reverse('position-index'))
        else:
            message = form.errors
            context = {
                'form': form,
                'segment': 'position',
                'group_segment': 'master',
                'crud': 'add',
                'message': message,
                'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
                'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='POSITION') if not request.user.is_superuser else Auth.objects.all(),
            }
            return render(request, 'home/position_add.html', context)
    else:
        form = FormPosition()
        context = {
            'form': form,
            'segment': 'position',
            'group_segment': 'master',
            'crud': 'add',
            'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
            'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='POSITION') if not request.user.is_superuser else Auth.objects.all(),
        }
        return render(request, 'home/position_add.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='POSITION')
def position_index(request):
    with connection.cursor() as cursor:
        cursor.execute("SELECT position_id, position_name FROM apps_position")
        positions = cursor.fetchall()

    context = {
        'data': positions,
        'attendance': attendance(request),
        'segment': 'position',
        'group_segment': 'master',
        'crud': 'index',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='POSITION') if not request.user.is_superuser else Auth.objects.all(),
    }

    return render(request, 'home/position_index.html', context)


# Update Position
@login_required(login_url='/login/')
@role_required(allowed_roles='POSITION')
def position_update(request, _id):
    positions = Position.objects.get(position_id=_id)
    if request.POST:
        form = FormPositionUpdate(
            request.POST, request.FILES, instance=positions)
        if form.is_valid():
            form.save()
            return HttpResponseRedirect(reverse('position-view', args=[_id, ]))
    else:
        form = FormPositionUpdate(instance=positions)

    message = form.errors
    context = {
        'form': form,
        'data': positions,
        'attendance': attendance(request),
        'segment': 'position',
        'group_segment': 'master',
        'crud': 'update',
        'message': message,
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='POSITION') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/position_view.html', context)


# Delete Position
@login_required(login_url='/login/')
@role_required(allowed_roles='POSITION')
def position_delete(request, _id):
    positions = Position.objects.get(position_id=_id)

    positions.delete()
    return HttpResponseRedirect(reverse('position-index'))


@login_required(login_url='/login/')
@role_required(allowed_roles='POSITION')
def position_view(request, _id):
    positions = Position.objects.get(position_id=_id)
    form = FormPositionView(instance=positions)

    context = {
        'form': form,
        'data': positions,
        'attendance': attendance(request),
        'segment': 'position',
        'group_segment': 'master',
        'crud': 'view',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='POSITION') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/position_view.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='MENU')
def menu_add(request):
    if request.POST:
        form = FormMenu(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return HttpResponseRedirect(reverse('menu-index'))
        else:
            message = form.errors
            context = {
                'form': form,
                'attendance': attendance(request),
                'segment': 'menu',
                'group_segment': 'master',
                'crud': 'add',
                'message': message,
                'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
                'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='MENU') if not request.user.is_superuser else Auth.objects.all(),
            }
            return render(request, 'home/menu_add.html', context)
    else:
        form = FormMenu()
        context = {
            'form': form,
            'attendance': attendance(request),
            'segment': 'menu',
            'group_segment': 'master',
            'crud': 'add',
            'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
            'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='MENU') if not request.user.is_superuser else Auth.objects.all(),
        }
        return render(request, 'home/menu_add.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='MENU')
def menu_index(request):
    with connection.cursor() as cursor:
        cursor.execute("SELECT menu_id, menu_name, menu_remark FROM apps_menu")
        menus = cursor.fetchall()

    context = {
        'data': menus,
        'attendance': attendance(request),
        'segment': 'menu',
        'group_segment': 'master',
        'crud': 'index',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='MENU') if not request.user.is_superuser else Auth.objects.all(),
    }

    return render(request, 'home/menu_index.html', context)


# Update Menu
@login_required(login_url='/login/')
@role_required(allowed_roles='MENU')
def menu_update(request, _id):
    menus = Menu.objects.get(menu_id=_id)
    if request.POST:
        form = FormMenuUpdate(request.POST, request.FILES, instance=menus)
        if form.is_valid():
            form.save()
            return HttpResponseRedirect(reverse('menu-view', args=[_id, ]))
    else:
        form = FormMenuUpdate(instance=menus)

    message = form.errors
    context = {
        'form': form,
        'data': menus,
        'attendance': attendance(request),
        'segment': 'menu',
        'group_segment': 'master',
        'crud': 'update',
        'message': message,
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='MENU') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/menu_view.html', context)


# Delete Menu
@login_required(login_url='/login/')
@role_required(allowed_roles='MENU')
def menu_delete(request, _id):
    menus = Menu.objects.get(menu_id=_id)

    menus.delete()
    return HttpResponseRedirect(reverse('menu-index'))


@login_required(login_url='/login/')
@role_required(allowed_roles='MENU')
def menu_view(request, _id):
    menus = Menu.objects.get(menu_id=_id)
    form = FormMenuView(instance=menus)

    context = {
        'form': form,
        'data': menus,
        'attendance': attendance(request),
        'segment': 'menu',
        'group_segment': 'master',
        'crud': 'view',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='MENU') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/menu_view.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='ATT-REPORT')
def report_attendance(request, _from_date, _to_date, _user):
    from_date = datetime.datetime.strptime(
        _from_date, '%Y-%m-%d').date() if _from_date != '0' else datetime.date.today()
    to_date = datetime.datetime.strptime(
        _to_date, '%Y-%m-%d').date() if _to_date != '0' else datetime.date.today()
    users = User.objects.all().exclude(is_superuser=True)

    if _user == 'all':
        attendances = Attendance.objects.filter(
            absence_date__gte=from_date, absence_date__lte=to_date)
    else:
        attendances = Attendance.objects.filter(
            user_id=_user, absence_date__gte=from_date, absence_date__lte=to_date)

    for att in attendances:
        att.lat_in = '{:.7f}'.format(att.lat_in) if att.lat_in else 0
        att.long_in = '{:.7f}'.format(att.long_in) if att.long_in else 0
        att.lat_out = '{:.7f}'.format(att.lat_out) if att.lat_out else 0
        att.long_out = '{:.7f}'.format(att.long_out) if att.long_out else 0

    context = {
        'data': attendances,
        'from_date': from_date,
        'to_date': to_date,
        'fromDate': _from_date,
        'toDate': _to_date,
        'selected_user': _user,
        'users': users,
        'attendance': attendance(request),
        'segment': 'att-report',
        'group_segment': 'report',
        'crud': 'index',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='ATT-REPORT') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/report_attendance.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='ATT-REPORT')
def report_attendance_toxl(request, _from_date, _to_date, _user):
    from_date = datetime.datetime.strptime(
        _from_date, '%Y-%m-%d').date() if _from_date != '0' else datetime.date.today()
    to_date = datetime.datetime.strptime(
        _to_date, '%Y-%m-%d').date() if _to_date != '0' else datetime.date.today()

    if _user == 'all':
        attendance = Attendance.objects.filter(
            absence_date__gte=from_date, absence_date__lte=to_date).annotate(
            lokasi_masuk=Concat(Value('https://www.google.com/maps/search/?api=1&query='), Cast('lat_in', CharField()), Value(
                ','), Cast('long_in', CharField()), output_field=CharField()),
            lokasi_pulang=Concat(Value('https://www.google.com/maps/search/?api=1&query='), Cast('lat_out', CharField()), Value(
                ','), Cast('long_out', CharField()), output_field=CharField()),
            pho_in=Concat(Value(host.url + 'apps/media/'),
                          'photo_in', output_field=CharField()),
            pho_out=Concat(Value(host.url + 'apps/media/'),
                           'photo_out', output_field=CharField())
        ).values_list(
            'user_id__username', 'absence_date', 'time_in', 'time_out', 'total_hours', 'lokasi_masuk', 'pho_in', 'lokasi_pulang', 'pho_out', 'status', 'note', 'address_in', 'address_out'
        )
    else:
        attendance = Attendance.objects.filter(
            user_id=_user, absence_date__gte=from_date, absence_date__lte=to_date).annotate(
            lokasi_masuk=Concat(Value('https://www.google.com/maps/search/?api=1&query='), Cast('lat_in', CharField()), Value(
                ','), Cast('long_in', CharField()), output_field=CharField()),
            lokasi_pulang=Concat(Value('https://www.google.com/maps/search/?api=1&query='), Cast('lat_out', CharField()), Value(
                ','), Cast('long_out', CharField()), output_field=CharField())
        ).values_list(
            'user_id__username', 'absence_date', 'time_in', 'time_out', 'total_hours', 'lokasi_masuk', 'photo_in', 'lokasi_pulang', 'photo_out', 'status', 'note', 'address_in', 'address_out'
        )

    # Create a HttpResponse object with the csv data
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    filename = 'laporan_kehadiran_' + \
        _from_date + '_' + '_to_' + _to_date + '_' + \
        '_user_' + _user + '.xlsx'
    response['Content-Disposition'] = 'attachment; filename=' + filename

    # Create an XlsxWriter workbook object and add a worksheet.
    workbook = xlsxwriter.Workbook(response, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Define column headers
    headers = ['Nama', 'Tanggal', 'Jam Masuk', 'Jam Pulang', 'Durasi Kerja', 'Lokasi Masuk', 'Foto Masuk',
               'Lokasi Pulang', 'Foto Pulang', 'Status', 'Keterangan']

    # Define cell formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#035a21',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
    })
    cell_format = workbook.add_format({'border': 1})
    date_format = workbook.add_format(
        {'border': 1, 'num_format': 'dd-mmm-yyyy', 'align': 'left'})
    time_format = workbook.add_format(
        {'border': 1, 'num_format': 'hh:mm', 'align': 'left'})
    late_time_format = workbook.add_format(
        {'border': 1, 'num_format': 'hh:mm', 'font_color': '#ea0606', 'align': 'left'})
    in_format = workbook.add_format({'border': 1, 'bg_color': '#a5d223'})
    out_format = workbook.add_format({'border': 1, 'bg_color': '#d1d1d1'})
    late_note_format = workbook.add_format(
        {'border': 1, 'font_color': 'white', 'bg_color': '#ea0606', 'align': 'left'})

    # Set column width
    worksheet.set_column('A:A', 25)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:E', 10)
    worksheet.set_column('F:F', 30)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 30)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:J', 10)
    worksheet.set_column('K:K', 15)

    # Write data to XlsxWriter Object
    for idx, record in enumerate(attendance):
        for col_idx, col_value in enumerate(record):
            if idx == 0:
                # Write the column headers on the first row
                if col_idx < len(headers):
                    worksheet.write(
                        idx, col_idx, headers[col_idx], header_format)
            # Write the data rows
            if col_idx == 1:
                worksheet.write(idx + 1, col_idx, col_value, date_format)
            elif col_idx == 2:
                if record[10] == 'Terlambat' and col_value:
                    worksheet.write(idx + 1, col_idx,
                                    col_value, late_time_format)
                else:
                    worksheet.write(idx + 1, col_idx, col_value, time_format)
            elif col_idx == 3:
                worksheet.write(idx + 1, col_idx, col_value, time_format)
            elif col_idx == 4:
                worksheet.write(idx + 1, col_idx, col_value, time_format)
            elif col_idx == 5:
                if col_value != 'https://www.google.com/maps/search/?api=1&query=,':
                    if record[11]:
                        worksheet.write_url(
                            idx + 1, col_idx, col_value, cell_format, record[11])
                    else:
                        worksheet.write_url(
                            idx + 1, col_idx, col_value, cell_format, 'Lihat Lokasi')
                else:
                    worksheet.write(idx + 1, col_idx, '', cell_format)
            elif col_idx == 6:
                if col_value != host.url + 'apps/media/':
                    worksheet.write_url(
                        idx + 1, col_idx, col_value, cell_format, 'Lihat Foto')
                else:
                    worksheet.write(idx + 1, col_idx, '', cell_format)
            elif col_idx == 7:
                if col_value != 'https://www.google.com/maps/search/?api=1&query=,':
                    if record[12]:
                        worksheet.write_url(
                            idx + 1, col_idx, col_value, cell_format, record[12])
                    else:
                        worksheet.write_url(
                            idx + 1, col_idx, col_value, cell_format, 'Lihat Lokasi')
                else:
                    worksheet.write(idx + 1, col_idx, '', cell_format)
            elif col_idx == 8:
                if col_value != host.url + 'apps/media/':
                    worksheet.write_url(
                        idx + 1, col_idx, col_value, cell_format, 'Lihat Foto')
                else:
                    worksheet.write(idx + 1, col_idx, '', cell_format)
            elif col_idx == 9:
                if col_value == 'Masuk':
                    worksheet.write(idx + 1, col_idx, col_value, in_format)
                elif col_value == 'Pulang':
                    worksheet.write(
                        idx + 1, col_idx, col_value, out_format)
                else:
                    worksheet.write(idx + 1, col_idx, col_value, cell_format)
            elif col_idx == 10:
                if col_value == 'Terlambat':
                    worksheet.write(idx + 1, col_idx,
                                    col_value, late_note_format)
                else:
                    worksheet.write(idx + 1, col_idx, col_value, cell_format)
            else:
                if col_idx < len(headers):
                    worksheet.write(idx + 1, col_idx, col_value, cell_format)

    # Close the workbook before sending the data.
    workbook.close()

    return response


@login_required(login_url='/login/')
@role_required(allowed_roles='CLOCK-IN')
def clock_in(request):
    hr = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Ahad']
    bln = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
           'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des']
    tgl = (hr[datetime.datetime.now().weekday()] + ', ' +
           datetime.datetime.now().strftime('%-d') + ' ' +
           bln[int(datetime.datetime.now().strftime('%m')) - 1] +
           ' ' + datetime.datetime.now().strftime('%Y'))

    context = {
        'attendance': attendance(request),
        'tgl': tgl,
        'segment': 'clock-in',
        'group_segment': 'attendance',
        'crud': 'detail',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='CLOCK-IN') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/clock_in.html', context)


@login_required(login_url='/login/')
@role_required(allowed_roles='CLOCK-OUT')
def clock_out(request):
    hr = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Ahad']
    bln = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
           'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des']
    tgl = (hr[datetime.datetime.now().weekday()] + ', ' +
           datetime.datetime.now().strftime('%-d') + ' ' +
           bln[int(datetime.datetime.now().strftime('%m')) - 1] +
           ' ' + datetime.datetime.now().strftime('%Y'))

    context = {
        'attendance': attendance(request),
        'tgl': tgl,
        'segment': 'clock-out',
        'group_segment': 'attendance',
        'crud': 'detail',
        'role': Auth.objects.filter(user_id=request.user.user_id).values_list('menu_id', flat=True),
        'btn': Auth.objects.get(user_id=request.user.user_id, menu_id='CLOCK-OUT') if not request.user.is_superuser else Auth.objects.all(),
    }
    return render(request, 'home/clock_out.html', context)


@login_required(login_url='/login/')
def submit_attendance(request):
    if request.POST:
        if request.POST.get('photo'):
            # save photo
            format, imgstr = request.POST.get('photo').split(';base64,')
            ext = format.split('/')[-1]
            img_data = base64.b64decode(imgstr)
            filename = f"{request.user.user_id}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.{ext}"
            file_path = os.path.join(
                settings.MEDIA_ROOT + '/attendance/', filename)
            with open(file_path, 'wb') as f:
                f.write(img_data)

            if not settings.DEBUG:
                # save to media
                file_path = '../../www/mahad/apps/media/' + 'attendance/' + filename
                with open(file_path, 'wb') as f:
                    f.write(img_data)

            geolocator = Nominatim(user_agent="myGeolocator")
            location = geolocator.reverse(
                f"{request.POST.get('latitude')}, {request.POST.get('longitude')}")

            if request.POST.get('status') == 'Masuk':
                Attendance.objects.update_or_create(
                    user_id=request.user.user_id,
                    absence_date=datetime.date.today(),
                    defaults={
                        'time_in': datetime.datetime.now(),
                        'lat_in': request.POST.get('latitude'),
                        'long_in': request.POST.get('longitude'),
                        'address_in': location.address if location else '',
                        'photo_in': 'attendance/' + filename,
                        'status': request.POST.get('status'),
                        'note': 'Terlambat' if datetime.datetime.now().time() > datetime.time(7, 15, 0) else '',
                    }
                )
            else:
                attendance = Attendance.objects.get(
                    user_id=request.user.user_id, absence_date=datetime.date.today())
                attendance.time_out = datetime.datetime.now()
                attendance.lat_out = request.POST.get('latitude')
                attendance.long_out = request.POST.get('longitude')
                attendance.address_out = location.address if location else ''
                attendance.photo_out = 'attendance/' + filename
                attendance.status = request.POST.get('status')
                attendance.save()

        return HttpResponseRedirect(reverse('home'))

    return render(request, 'home/home.html')


@login_required(login_url='/login/')
@role_required(allowed_roles='CLOCK-IN')
def clock_in_success(request):
    return render(request, 'home/clock_in_success.html')


@login_required(login_url='/login/')
@role_required(allowed_roles='CLOCK-OUT')
def clock_out_success(request):
    return render(request, 'home/clock_out_success.html')
