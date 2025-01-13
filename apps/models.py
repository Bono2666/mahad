from django.contrib.auth.models import AbstractUser
from django.db import models
from django.utils import timezone
from crum import get_current_user
from decimal import Decimal
from re import sub
from django.db import models
from tinymce.models import HTMLField


class User(AbstractUser):
    is_active = models.BooleanField(default=True)
    user_id = models.CharField(max_length=50, primary_key=True)
    username = models.CharField(max_length=50)
    position = models.ForeignKey(
        'Position', on_delete=models.CASCADE, null=True)
    signature = models.ImageField(upload_to='signature/', null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(User, self).save(*args, **kwargs)

    USERNAME_FIELD = 'user_id'
    REQUIRED_FIELDS = ['username']

    def __str__(self):
        return self.username


class Room(models.Model):
    room_id = models.CharField(max_length=50, primary_key=True)
    room_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.room_id = self.room_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Room, self).save(*args, **kwargs)

    def __str__(self):
        return self.room_name


class Task(models.Model):
    task_id = models.BigAutoField(primary_key=True)
    task_name = models.CharField(max_length=50)
    room = models.ForeignKey(Room, on_delete=models.CASCADE)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Task, self).save(*args, **kwargs)

    def __str__(self):
        return self.task_name


class Checklist(models.Model):
    checklist_id = models.BigAutoField(primary_key=True)
    room = models.ForeignKey(Room, on_delete=models.CASCADE, null=True)
    task = models.ForeignKey(Task, on_delete=models.CASCADE)
    checklist_status = models.CharField(
        max_length=20, default='Belum Dikerjakan')
    checklist_remark = models.CharField(max_length=200, null=True)
    checklist_attachment = models.ImageField(upload_to='checklist/', null=True)
    checklist_start = models.DateTimeField(null=True)
    checklist_end = models.DateTimeField(null=True)
    checklist_date = models.DateTimeField(null=True)
    checklist_duration = models.IntegerField(default=0)
    checklist_by = models.ForeignKey(
        User, on_delete=models.CASCADE, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.checklist_date:
            self.checklist_date = timezone.now()
        if self.checklist_start:
            self.checklist_by = get_current_user()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Checklist, self).save(*args, **kwargs)

    def __str__(self):
        return self.checklist_name


class Distributor(models.Model):
    distributor_id = models.CharField(max_length=50, primary_key=True)
    distributor_name = models.CharField(max_length=50)
    sap_code = models.CharField(max_length=50, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.distributor_id = self.distributor_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Distributor, self).save(*args, **kwargs)

    def __str__(self):
        return self.distributor_name


class Channel(models.Model):
    channel_id = models.CharField(max_length=5, primary_key=True)
    channel_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.channel_id = self.channel_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Channel, self).save(*args, **kwargs)

    def __str__(self):
        return self.channel_id + ' - ' + self.channel_name


class Equipment(models.Model):
    equipment_id = models.CharField(max_length=50, primary_key=True)
    equipment_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.equipment_id = self.equipment_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Equipment, self).save(*args, **kwargs)

    def __str__(self):
        return self.equipment_name


class Category(models.Model):
    category_id = models.CharField(max_length=50, primary_key=True)
    category_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.category_id = self.category_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Category, self).save(*args, **kwargs)

    def __str__(self):
        return self.category_name


class Package(models.Model):
    package_id = models.CharField(max_length=50, primary_key=True)
    package_name = models.CharField(max_length=50)
    category = models.ForeignKey(Category, on_delete=models.CASCADE)
    male_price = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    female_price = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    box = models.IntegerField(default=0)
    quantity = models.IntegerField(default=0)
    type = models.CharField(max_length=10, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        self.package_id = self.package_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Package, self).save(*args, **kwargs)

    def __str__(self):
        return self.package_name


class Bag(models.Model):
    package = models.ForeignKey(Package, on_delete=models.CASCADE)
    equipment = models.ForeignKey(Equipment, on_delete=models.CASCADE)
    extra_price = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    default = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['package', 'equipment'], name='unique_bag')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Bag, self).save(*args, **kwargs)


class Pack(models.Model):
    package = models.ForeignKey(Package, on_delete=models.CASCADE)
    equipment = models.ForeignKey(Equipment, on_delete=models.CASCADE)
    extra_price = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    default = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['package', 'equipment'], name='unique_pack')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Pack, self).save(*args, **kwargs)


class Addon(models.Model):
    package = models.ForeignKey(Package, on_delete=models.CASCADE)
    equipment = models.ForeignKey(Equipment, on_delete=models.CASCADE)
    extra_price = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    default = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['package', 'equipment'], name='unique_addon')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(Addon, self).save(*args, **kwargs)


class AreaSales(models.Model):
    area_id = models.CharField(max_length=50, primary_key=True)
    area_name = models.CharField(max_length=50)
    manager = models.CharField(max_length=50)
    bank_account = models.CharField(max_length=200, null=True)
    address = models.CharField(max_length=200, null=True)
    district = models.CharField(max_length=50, null=True)
    city = models.CharField(max_length=50, null=True)
    postal_code = models.CharField(max_length=10, null=True)
    form = models.CharField(max_length=200, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.area_id = self.area_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(AreaSales, self).save(*args, **kwargs)

    def __str__(self):
        return self.area_name

    def get_area_sales_children(self):
        return self.areasalesdetail_set.all()


class AreaChannel(models.Model):
    area_id = models.CharField(max_length=50, primary_key=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(AreaChannel, self).save(*args, **kwargs)


class AreaSalesDetail(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    distributor = models.ForeignKey(Distributor, on_delete=models.CASCADE)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'distributor'], name='unique_area_distributor')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(AreaSalesDetail, self).save(*args, **kwargs)

    def __str__(self):
        return self.distributor_name


class AreaChannelDetail(models.Model):
    area = models.ForeignKey(AreaChannel, on_delete=models.CASCADE)
    channel = models.ForeignKey(Channel, on_delete=models.CASCADE)
    status = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'channel'], name='unique_area_channel')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(AreaChannelDetail, self).save(*args, **kwargs)

    def __str__(self):
        return self.area_id


class AreaUser(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'user'], name='unique_area_user')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(AreaUser, self).save(*args, **kwargs)

    def __str__(self):
        return self.area_id


class Position(models.Model):
    position_id = models.CharField(
        max_length=3, primary_key=True, help_text='Max 3 digits Position shortname.')
    position_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.position_id = self.position_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Position, self).save(*args, **kwargs)

    def __str__(self):
        return self.position_name


class Menu(models.Model):
    menu_id = models.CharField(max_length=50, primary_key=True)
    menu_name = models.CharField(max_length=50)
    menu_remark = models.CharField(max_length=200, null=True, blank=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.menu_id = self.menu_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Menu, self).save(*args, **kwargs)

    def __str__(self):
        return self.menu_name


class Auth(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    menu = models.ForeignKey(Menu, on_delete=models.CASCADE)
    add = models.BooleanField(default=False)
    edit = models.BooleanField(default=False)
    delete = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['user', 'menu'], name='unique_user_menu')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Auth, self).save(*args, **kwargs)

    def __str__(self):
        return self.menu.menu_name


class Budget(models.Model):
    budget_id = models.CharField(max_length=50, primary_key=True)
    budget_year = models.CharField(max_length=4)
    budget_month = models.CharField(max_length=2)
    budget_area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    budget_distributor = models.ForeignKey(
        Distributor, on_delete=models.CASCADE)
    budget_amount = models.DecimalField(
        max_digits=12, decimal_places=0)
    budget_upping = models.DecimalField(
        max_digits=12, decimal_places=0)
    budget_total = models.DecimalField(
        max_digits=12, decimal_places=0)
    budget_balance = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_status = models.CharField(max_length=15)
    budget_new = models.BooleanField(default=False)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.budget_id:
            self.budget_id = 'UBS/' + self.budget_area.area_id + '/' + \
                self.budget_distributor.distributor_id + '/' + \
                self.budget_month + '/' + self.budget_year
        self.budget_amount = Decimal(
            sub(r'[^\d.]', '', str(self.budget_amount)))
        self.budget_upping = Decimal(
            sub(r'[^\d.]', '', str(self.budget_upping)))
        self.budget_total = self.budget_amount + self.budget_upping
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Budget, self).save(*args, **kwargs)

    def __str__(self):
        return self.budget_id


class BudgetDetail(models.Model):
    budget = models.ForeignKey(Budget, on_delete=models.CASCADE)
    budget_channel = models.ForeignKey(Channel, on_delete=models.CASCADE)
    budget_percent = models.DecimalField(
        max_digits=3, decimal_places=0, default=0)
    budget_amount = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_upping = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_total = models.DecimalField(max_digits=12, decimal_places=0)
    budget_proposed = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_program = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_claim = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    budget_balance = models.DecimalField(max_digits=12, decimal_places=0)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['budget', 'budget_channel'], name='unique_budget_channel')
        ]

    def save(self, *args, **kwargs):
        self.budget_total = self.budget_amount + self.budget_upping
        self.budget_balance = self.budget_total - self.budget_proposed
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(BudgetDetail, self).save(*args, **kwargs)


class UploadLog(models.Model):
    document = models.CharField(max_length=50)
    document_id = models.CharField(max_length=50)
    description = models.CharField(max_length=200)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(UploadLog, self).save(*args, **kwargs)


class BudgetRelease(models.Model):
    budget = models.ForeignKey(Budget, on_delete=models.CASCADE)
    budget_approval_id = models.CharField(max_length=50, null=True)
    budget_approval_name = models.CharField(max_length=50, null=True)
    budget_approval_email = models.CharField(max_length=50, null=True)
    budget_approval_position = models.CharField(max_length=50, null=True)
    budget_approval_date = models.DateTimeField(null=True)
    budget_approval_status = models.CharField(max_length=1, default='N')
    upping_note = models.CharField(max_length=200, null=True)
    percentage_note = models.CharField(max_length=200, null=True)
    return_note = models.CharField(max_length=200, null=True)
    sequence = models.IntegerField(default=0)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['budget', 'budget_approval_id'], name='unique_budget_approval')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(BudgetRelease, self).save(*args, **kwargs)


class ProposalRelease(models.Model):
    proposal = models.ForeignKey('Proposal', on_delete=models.CASCADE)
    proposal_approval_id = models.CharField(max_length=50, null=True)
    proposal_approval_name = models.CharField(max_length=50, null=True)
    proposal_approval_email = models.CharField(max_length=50, null=True)
    proposal_approval_position = models.CharField(max_length=50, null=True)
    proposal_approval_date = models.DateTimeField(null=True)
    proposal_approval_status = models.CharField(max_length=1, default='N')
    sequence = models.IntegerField(default=0)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    return_to = models.BooleanField(default=False)
    revise_note = models.CharField(max_length=200, null=True)
    return_note = models.CharField(max_length=200, null=True)
    reject_note = models.CharField(max_length=200, null=True)
    mail_sent = models.BooleanField(default=False)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['proposal', 'proposal_approval_id'], name='unique_proposal_approval')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(ProposalRelease, self).save(*args, **kwargs)


class BudgetApproval(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    approver = models.ForeignKey(User, on_delete=models.CASCADE)
    sequence = models.IntegerField(default=0)
    position = models.ForeignKey(Position, on_delete=models.CASCADE, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'approver'], name='unique_area_approver')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(BudgetApproval, self).save(*args, **kwargs)

    def __str__(self):
        return self.area.area_name


class ProposalMatrix(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    channel = models.ForeignKey(Channel, on_delete=models.CASCADE, null=True)
    approver = models.ForeignKey(User, on_delete=models.CASCADE)
    sequence = models.IntegerField(default=0)
    position = models.ForeignKey(Position, on_delete=models.CASCADE, null=True)
    return_to = models.BooleanField(default=False)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'channel', 'approver'], name='unique_proposal_approver')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(ProposalMatrix, self).save(*args, **kwargs)

    def __str__(self):
        return self.area.area_name


class Closing(models.Model):
    document = models.CharField(max_length=50, primary_key=True)
    year_closed = models.CharField(max_length=4)
    month_closed = models.CharField(max_length=2)
    year_open = models.CharField(max_length=4)
    month_open = models.CharField(max_length=2)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    def save(self, *args, **kwargs):
        self.document = self.document.upper().replace(' ', '_')
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Closing, self).save(*args, **kwargs)


class Division(models.Model):
    division_id = models.BigAutoField(primary_key=True)
    division_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Division, self).save(*args, **kwargs)

    def __str__(self):
        return self.division_name


class Proposal(models.Model):
    proposal_id = models.CharField(max_length=50, primary_key=True)
    proposal_date = models.DateTimeField(null=True)
    type = models.CharField(max_length=1, default='B')
    budget = models.ForeignKey(Budget, on_delete=models.CASCADE)
    channel = models.CharField(max_length=5)
    division = models.ForeignKey(Division, on_delete=models.CASCADE)
    program_name = models.CharField(max_length=200)
    products = models.TextField()
    area = models.CharField(max_length=50)
    period_start = models.DateTimeField(null=True)
    period_end = models.DateTimeField(null=True)
    duration = models.IntegerField(default=0)
    background = models.TextField(null=True)
    objectives = models.TextField()
    mechanism = models.TextField()
    remarks = models.CharField(max_length=200, null=True)
    attachment = models.FileField(upload_to='proposal/', null=True)
    total_cost = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    proposal_claim = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    balance = models.DecimalField(max_digits=12, decimal_places=0, default=0)
    roi = models.DecimalField(
        max_digits=10, decimal_places=0, default=0)
    status = models.CharField(max_length=15, default='DRAFT')
    seq_number = models.IntegerField(default=0)
    entry_pos = models.CharField(max_length=5, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        self.duration = (self.period_end - self.period_start).days + 1
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(Proposal, self).save(*args, **kwargs)


class IncrementalSales(models.Model):
    proposal = models.ForeignKey(Proposal, on_delete=models.CASCADE)
    product = models.CharField(max_length=50)
    swop_carton = models.IntegerField(default=0)
    swop_nom_carton = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    swop_nom = models.DecimalField(
        max_digits=14, decimal_places=0, default=0)
    swp_carton = models.IntegerField(default=0)
    swp_nom_carton = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    swp_nom = models.DecimalField(
        max_digits=14, decimal_places=0, default=0)
    incrp_carton = models.IntegerField(default=0)
    incrp_nom = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    incpst_carton = models.DecimalField(
        max_digits=5, decimal_places=1, default=0)
    incpst_nom = models.DecimalField(max_digits=5, decimal_places=1, default=0)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['proposal', 'product'], name='unique_proposal_product')
        ]

    def save(self, *args, **kwargs):
        self.incrp_carton = self.swp_carton - self.swop_carton
        self.incrp_nom = self.swp_nom - self.swop_nom
        self.incpst_carton = (self.incrp_carton /
                              self.swop_carton) * 100 if self.swop_carton > 0 else 0
        self.incpst_nom = (self.incrp_nom / self.swop_nom) * \
            100 if self.swop_nom > 0 else 0
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(IncrementalSales, self).save(*args, **kwargs)


class ProjectedCost(models.Model):
    proposal = models.ForeignKey(Proposal, on_delete=models.CASCADE)
    activities = models.CharField(max_length=200)
    cost = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['proposal', 'activities'], name='unique_proposal_activities')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(ProjectedCost, self).save(*args, **kwargs)


class Program(models.Model):
    program_id = models.CharField(max_length=50, primary_key=True)
    program_date = models.DateTimeField(null=True)
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE, null=True)
    proposal = models.ForeignKey(Proposal, on_delete=models.CASCADE, null=True)
    deadline = models.DateTimeField(null=True)
    content = HTMLField()
    approval = HTMLField(null=True)
    seq_number = models.IntegerField(default=0)
    status = models.CharField(max_length=15, default='PENDING')
    entry_pos = models.CharField(max_length=5, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(Program, self).save(*args, **kwargs)


class ProgramMatrix(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    channel = models.ForeignKey(Channel, on_delete=models.CASCADE, null=True)
    approver = models.ForeignKey(User, on_delete=models.CASCADE)
    sequence = models.IntegerField(default=0)
    position = models.ForeignKey(Position, on_delete=models.CASCADE, null=True)
    return_to = models.BooleanField(default=False)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'channel', 'approver'], name='unique_program_approver')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(ProgramMatrix, self).save(*args, **kwargs)

    def __str__(self):
        return self.area.area_name


class ProgramRelease(models.Model):
    program = models.ForeignKey('Program', on_delete=models.CASCADE)
    program_approval_id = models.CharField(max_length=50, null=True)
    program_approval_name = models.CharField(max_length=50, null=True)
    program_approval_email = models.CharField(max_length=50, null=True)
    program_approval_position = models.CharField(max_length=50, null=True)
    program_approval_date = models.DateTimeField(null=True)
    program_approval_status = models.CharField(max_length=1, default='N')
    sequence = models.IntegerField(default=0)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    return_to = models.BooleanField(default=False)
    revise_note = models.CharField(max_length=200, null=True)
    return_note = models.CharField(max_length=200, null=True)
    reject_note = models.CharField(max_length=200, null=True)
    mail_sent = models.BooleanField(default=False)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True)
    update_by = models.CharField(max_length=50, null=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['program', 'program_approval_id'], name='unique_program_approval')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(ProgramRelease, self).save(*args, **kwargs)


class Claim(models.Model):
    claim_id = models.CharField(max_length=50, primary_key=True)
    claim_date = models.DateTimeField(null=True)
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE, null=True)
    proposal = models.ForeignKey(Proposal, on_delete=models.CASCADE, null=True)
    additional_proposal = models.CharField(max_length=50, null=True)
    program = models.ForeignKey(Program, on_delete=models.CASCADE, null=True)
    invoice = models.CharField(max_length=50, null=True)
    invoice_date = models.DateTimeField(null=True)
    due_date = models.DateTimeField(null=True)
    amount = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    additional_amount = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    tax = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    total = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    total_claim = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    remarks = models.TextField(null=True)
    status = models.CharField(max_length=15, default='PENDING')
    seq_number = models.IntegerField(default=0)
    entry_pos = models.CharField(max_length=5, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        self.tax = self.total_claim * Decimal(0.11)
        self.total = self.total_claim + self.tax
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(Claim, self).save(*args, **kwargs)


class ClaimMatrix(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    channel = models.ForeignKey(Channel, on_delete=models.CASCADE, null=True)
    approver = models.ForeignKey(User, on_delete=models.CASCADE)
    sequence = models.IntegerField(default=0)
    position = models.ForeignKey(Position, on_delete=models.CASCADE, null=True)
    return_to = models.BooleanField(default=False)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'channel', 'approver'], name='unique_claim_approver')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(ClaimMatrix, self).save(*args, **kwargs)


class ClaimRelease(models.Model):
    claim = models.ForeignKey('Claim', on_delete=models.CASCADE)
    claim_approval_id = models.CharField(max_length=50, null=True)
    claim_approval_name = models.CharField(max_length=50, null=True)
    claim_approval_email = models.CharField(max_length=50, null=True)
    claim_approval_position = models.CharField(max_length=50, null=True)
    claim_approval_date = models.DateTimeField(null=True)
    claim_approval_status = models.CharField(max_length=1, default='N')
    sequence = models.IntegerField(default=0)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    return_to = models.BooleanField(default=False)
    revise_note = models.CharField(max_length=200, null=True)
    return_note = models.CharField(max_length=200, null=True)
    reject_note = models.CharField(max_length=200, null=True)
    mail_sent = models.BooleanField(default=False)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['claim', 'claim_approval_id'], name='unique_claim_approval')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(ClaimRelease, self).save(*args, **kwargs)


class CL(models.Model):
    cl_id = models.CharField(max_length=50, primary_key=True)
    cl_date = models.DateTimeField(null=True)
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE, null=True)
    distributor = models.ForeignKey(
        Distributor, on_delete=models.CASCADE, null=True)
    status = models.CharField(max_length=15, default='DRAFT')
    seq_number = models.IntegerField(default=0)
    entry_pos = models.CharField(max_length=5, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(CL, self).save(*args, **kwargs)


class CLDetail(models.Model):
    cl_id = models.ForeignKey(CL, on_delete=models.CASCADE)
    claim = models.ForeignKey('Claim', on_delete=models.CASCADE)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(null=True, blank=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['claim', 'cl_id'], name='unique_claim_list')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(CLDetail, self).save(*args, **kwargs)


class CLMatrix(models.Model):
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    approver = models.ForeignKey(User, on_delete=models.CASCADE)
    sequence = models.IntegerField(default=0)
    position = models.ForeignKey(Position, on_delete=models.CASCADE, null=True)
    return_to = models.BooleanField(default=False)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['area', 'approver'], name='unique_cl_approver')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.update_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(CLMatrix, self).save(*args, **kwargs)


class CLRelease(models.Model):
    cl_id = models.ForeignKey(CL, on_delete=models.CASCADE)
    cl_approval_id = models.CharField(max_length=50, null=True)
    cl_approval_name = models.CharField(max_length=50, null=True)
    cl_approval_email = models.CharField(max_length=50, null=True)
    cl_approval_position = models.CharField(max_length=50, null=True)
    cl_approval_date = models.DateTimeField(null=True)
    cl_approval_status = models.CharField(max_length=1, default='N')
    sequence = models.IntegerField(default=0)
    limit = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    return_to = models.BooleanField(default=False)
    revise_note = models.CharField(max_length=200, null=True)
    return_note = models.CharField(max_length=200, null=True)
    reject_note = models.CharField(max_length=200, null=True)
    mail_sent = models.BooleanField(default=False)
    approve = models.BooleanField(default=False)
    revise = models.BooleanField(default=False)
    returned = models.BooleanField(default=False)
    reject = models.BooleanField(default=False)
    printed = models.BooleanField(default=False)
    notif = models.BooleanField(default=False)
    as_approved = models.CharField(max_length=15, null=True)
    entry_date = models.DateTimeField(
        null=True, blank=True, auto_now_add=True)
    entry_by = models.CharField(max_length=50, null=True, blank=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['cl_id', 'cl_approval_id'], name='unique_cl_approval')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.update_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(CLRelease, self).save(*args, **kwargs)


class Region(models.Model):
    region_id = models.CharField(max_length=50, primary_key=True)
    region_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        self.region_id = self.region_id.upper()
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(Region, self).save(*args, **kwargs)

    def __str__(self):
        return self.region_name


class RegionDetail(models.Model):
    region = models.ForeignKey(Region, on_delete=models.CASCADE)
    area = models.ForeignKey(AreaSales, on_delete=models.CASCADE)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['region', 'area'], name='unique_region_area')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.update_by = get_current_user().username
        self.update_date = timezone.now()
        self.update_by = get_current_user().username
        super(RegionDetail, self).save(*args, **kwargs)


class Customer(models.Model):
    customer_id = models.BigAutoField(primary_key=True)
    customer_name = models.CharField(max_length=200)
    customer_address = models.CharField(max_length=200, null=True)
    customer_district = models.CharField(max_length=50, null=True)
    customer_city = models.CharField(max_length=50, null=True)
    customer_province = models.CharField(max_length=50, null=True)
    customer_phone = models.CharField(max_length=50, null=True)
    customer_phone2 = models.CharField(max_length=50, null=True)
    customer_email = models.CharField(max_length=50, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Customer, self).save(*args, **kwargs)


class CustomerDetail(models.Model):
    customer = models.ForeignKey(Customer, on_delete=models.CASCADE)
    child_name = models.CharField(max_length=200)
    child_birth = models.DateField(null=True)
    child_sex = models.CharField(max_length=1, null=True)
    child_father = models.CharField(max_length=200, null=True)
    child_mother = models.CharField(max_length=200, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['customer', 'child_name'], name='unique_customer_child')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(CustomerDetail, self).save(*args, **kwargs)


class BoxType(models.Model):
    box_type_id = models.BigAutoField(primary_key=True)
    box_type_name = models.CharField(max_length=50)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(BoxType, self).save(*args, **kwargs)

    def __str__(self):
        return self.box_type_name


class Souvenir(models.Model):
    souvenir_id = models.BigAutoField(primary_key=True)
    souvenir_name = models.CharField(max_length=200)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(Souvenir, self).save(*args, **kwargs)

    def __str__(self):
        return self.souvenir_name


class Order(models.Model):
    order_id = models.CharField(max_length=50, primary_key=True)
    regional = models.ForeignKey(
        AreaSales, on_delete=models.CASCADE, null=True)
    order_date = models.DateTimeField(null=True)
    customer_name = models.CharField(max_length=200)
    customer_phone = models.CharField(max_length=50, null=True)
    customer_phone2 = models.CharField(max_length=50, null=True)
    customer_email = models.CharField(max_length=50, null=True)
    customer_address = models.CharField(max_length=200, null=True)
    customer_district = models.CharField(max_length=50, null=True)
    customer_city = models.CharField(max_length=50, null=True)
    customer_province = models.CharField(max_length=50, null=True)
    delivery_date = models.DateTimeField(null=True)
    time_arrival = models.CharField(max_length=50, null=True)
    total_order = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    down_payment = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    discount = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    pending_payment = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    use_photo = models.BooleanField(default=False)
    witnessed = models.BooleanField(default=False)
    info_source = models.CharField(max_length=50, null=True)
    souvenir = models.ForeignKey(Souvenir, on_delete=models.CASCADE, null=True)
    order_note = models.CharField(max_length=200, null=True)
    cs = models.CharField(max_length=50, null=True)
    order_status = models.CharField(max_length=15, default='PENDING')
    seq_number = models.IntegerField(default=0)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        self.pending_payment = self.total_order - self.down_payment - self.discount
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = 'customer'
        self.update_date = timezone.now()
        self.update_by = 'customer'
        super(Order, self).save(*args, **kwargs)


class OrderChild(models.Model):
    order = models.ForeignKey(Order, on_delete=models.CASCADE)
    child_name = models.CharField(max_length=200)
    child_birth = models.DateField(null=True)
    child_sex = models.CharField(max_length=1, null=True)
    child_father = models.CharField(max_length=200, null=True)
    child_mother = models.CharField(max_length=200, null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['order', 'child_name'], name='unique_order_child')
        ]

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = 'customer'
        self.update_date = timezone.now()
        self.update_by = 'customer'
        super(OrderChild, self).save(*args, **kwargs)


class CashIn(models.Model):
    cashin_id = models.BigAutoField(primary_key=True)
    order = models.ForeignKey(Order, on_delete=models.CASCADE)
    cashin_date = models.DateTimeField(null=True)
    cashin_type = models.CharField(max_length=50, null=True)
    cashin_amount = models.DecimalField(
        max_digits=12, decimal_places=0, default=0)
    cashin_note = models.CharField(max_length=200, null=True)
    bank = models.CharField(max_length=50, null=True)
    evidence = models.FileField(upload_to='cashin/', null=True)
    entry_date = models.DateTimeField(null=True)
    entry_by = models.CharField(max_length=50, null=True)
    update_date = models.DateTimeField(
        null=True, blank=True, auto_now=True)
    update_by = models.CharField(max_length=50, null=True, blank=True)

    def save(self, *args, **kwargs):
        if not self.entry_date:
            self.entry_date = timezone.now()
            self.entry_by = get_current_user().user_id
        self.update_date = timezone.now()
        self.update_by = get_current_user().user_id
        super(CashIn, self).save(*args, **kwargs)
