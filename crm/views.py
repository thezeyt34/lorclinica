import os
from datetime import timedelta
from decimal import Decimal
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login, logout, authenticate, update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import PasswordChangeForm
from django.http import HttpResponse, HttpResponseForbidden
from django.utils import timezone
from django.db.models import Sum, Count, F, Q
from django.db.models.functions import TruncDay, TruncWeek, TruncMonth
from django.contrib import messages
from django.core.paginator import Paginator

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from crm.models import User, Kirim, Chiqim

# Load Cyrillic font for PDF
font_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
arial_path = os.path.join(font_dir, 'arial.ttf')
arial_bold_path = os.path.join(font_dir, 'arialbd.ttf')

CYRILLIC_FONT_SUPPORTED = False
if os.path.exists(arial_path) and os.path.exists(arial_bold_path):
    try:
        pdfmetrics.registerFont(TTFont('Arial', arial_path))
        pdfmetrics.registerFont(TTFont('Arial-Bold', arial_bold_path))
        CYRILLIC_FONT_SUPPORTED = True
    except Exception:
        pass

def admin_required(view_func):
    def _wrapped_view(request, *args, **kwargs):
        if request.user.is_authenticated and request.user.is_admin():
            return view_func(request, *args, **kwargs)
        messages.error(request, "Доступ запрещен. Требуются права администратора.")
        return redirect('kirim')
    return _wrapped_view

def login_view(request):
    if request.user.is_authenticated:
        if request.user.is_admin():
            return redirect('dashboard')
        return redirect('kirim')
        
    if request.method == 'POST':
        username = request.POST.get('username')
        passw = request.POST.get('password')
        user = authenticate(request, username=username, password=passw)
        if user is not None:
            login(request, user)
            if user.is_admin():
                return redirect('dashboard')
            return redirect('kirim')
        else:
            messages.error(request, "Неверное имя пользователя или пароль.")
            
    return render(request, 'crm/login.html')

def logout_view(request):
    logout(request)
    return redirect('login')

def get_date_range(period):
    now = timezone.now()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    
    if period == 'today':
        return today_start, now
    elif period == 'week':
        return now - timedelta(days=7), now
    elif period == 'month':
        return now - timedelta(days=30), now
    elif period == '3months':
        return now - timedelta(days=90), now
    elif period == 'year':
        return now - timedelta(days=365), now
    else: # all time
        return None, None

@login_required
@admin_required
def dashboard_view(request):
    now = timezone.now()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    year_start = now.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)

    # 1. Earned today (Сегодня заработано)
    earned_today = Kirim.objects.filter(created_at__gte=today_start).aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    # 2. Spent today (Сегодня потрачено)
    spent_today = Chiqim.objects.filter(created_at__gte=today_start).aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    # 3. Today profit (Прибыль сегодня)
    profit_today = earned_today - spent_today
    # 4. Month income (Доход за месяц)
    income_month = Kirim.objects.filter(created_at__gte=month_start).aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    # 5. Year income (Доход за год)
    income_year = Kirim.objects.filter(created_at__gte=year_start).aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    # 6. Clients count (Количество клиентов)
    clients_count = Kirim.objects.values('client_name').distinct().count()
    # 7. Purchased items count (Количество купленных товаров - сумма quantites)
    purchased_items = Chiqim.objects.aggregate(sum=Sum('quantity'))['sum'] or 0
    # 8. Cash register balance (Остаток кассы)
    total_kirim = Kirim.objects.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    total_chiqim = Chiqim.objects.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    cash_balance = total_kirim - total_chiqim

    # Top Clients (Топ клиентов)
    top_clients = Kirim.objects.values('client_name').annotate(total_paid=Sum('amount')).order_by('-total_paid')[:5]

    # Action History (История действий - последние 10)
    recent_kirims = Kirim.objects.all().order_by('-created_at')[:10]
    recent_chiqims = Chiqim.objects.all().order_by('-created_at')[:10]
    
    action_history = []
    for k in recent_kirims:
        action_history.append({
            'type': 'kirim',
            'author': k.author.username,
            'client_or_item': k.client_name,
            'amount': k.amount,
            'created_at': k.created_at
        })
    for c in recent_chiqims:
        action_history.append({
            'type': 'chiqim',
            'author': c.author.username,
            'client_or_item': f"{c.item_name} ({c.quantity} шт.)",
            'amount': c.total_price,
            'created_at': c.created_at
        })
    action_history = sorted(action_history, key=lambda x: x['created_at'], reverse=True)[:10]

    # Chart.js metrics
    # Daily profits for last 30 days
    last_30_days = now - timedelta(days=30)
    daily_kirim = Kirim.objects.filter(created_at__gte=last_30_days).annotate(day=TruncDay('created_at')).values('day').annotate(total=Sum('amount'))
    daily_chiqim = Chiqim.objects.filter(created_at__gte=last_30_days).annotate(day=TruncDay('created_at')).values('day').annotate(total=Sum('total_price'))
    
    daily_data = {}
    for i in range(31):
        d = (now - timedelta(days=i)).date()
        daily_data[d] = {'kirim': Decimal(0), 'chiqim': Decimal(0)}
    
    for dk in daily_kirim:
        daily_data[dk['day'].date()]['kirim'] = dk['total'] or Decimal(0)
    for dc in daily_chiqim:
        daily_data[dc['day'].date()]['chiqim'] = dc['total'] or Decimal(0)
        
    sorted_days = sorted(daily_data.keys())
    chart_days_labels = [d.strftime('%d.%m') for d in sorted_days]
    chart_days_profits = [float(daily_data[d]['kirim'] - daily_data[d]['chiqim']) for d in sorted_days]
    chart_days_income = [float(daily_data[d]['kirim']) for d in sorted_days]

    # Weekly profits for last 8 weeks
    last_8_weeks = now - timedelta(weeks=8)
    weekly_kirim = Kirim.objects.filter(created_at__gte=last_8_weeks).annotate(week=TruncWeek('created_at')).values('week').annotate(total=Sum('amount'))
    weekly_chiqim = Chiqim.objects.filter(created_at__gte=last_8_weeks).annotate(week=TruncWeek('created_at')).values('week').annotate(total=Sum('total_price'))
    
    weekly_data = {}
    for i in range(8):
        w = (now - timedelta(weeks=i)).date() - timedelta(days=(now - timedelta(weeks=i)).weekday())
        weekly_data[w] = {'kirim': Decimal(0), 'chiqim': Decimal(0)}
        
    for wk in weekly_kirim:
        weekly_data[wk['week'].date()]['kirim'] = wk['total'] or Decimal(0)
    for wc in weekly_chiqim:
        weekly_data[wc['week'].date()]['chiqim'] = wc['total'] or Decimal(0)
        
    sorted_weeks = sorted(weekly_data.keys())
    chart_weeks_labels = [f"Нед {w.strftime('%d.%m')}" for w in sorted_weeks]
    chart_weeks_profits = [float(weekly_data[w]['kirim'] - weekly_data[w]['chiqim']) for w in sorted_weeks]

    # Monthly profits for last 6 months
    last_6_months = now - timedelta(days=180)
    monthly_kirim = Kirim.objects.filter(created_at__gte=last_6_months).annotate(month=TruncMonth('created_at')).values('month').annotate(total=Sum('amount'))
    monthly_chiqim = Chiqim.objects.filter(created_at__gte=last_6_months).annotate(month=TruncMonth('created_at')).values('month').annotate(total=Sum('total_price'))
    
    monthly_data = {}
    for i in range(6):
        # subtract months safely
        d = (now - timedelta(days=30*i)).replace(day=1).date()
        monthly_data[d] = {'kirim': Decimal(0), 'chiqim': Decimal(0)}
        
    for mk in monthly_kirim:
        monthly_data[mk['month'].date()]['kirim'] = mk['total'] or Decimal(0)
    for mc in monthly_chiqim:
        monthly_data[mc['month'].date()]['chiqim'] = mc['total'] or Decimal(0)
        
    sorted_months = sorted(monthly_data.keys())
    # Russian month names map
    ru_months = {1: 'Янв', 2: 'Фев', 3: 'Мар', 4: 'Апр', 5: 'Май', 6: 'Июн', 7: 'Июл', 8: 'Авг', 9: 'Сен', 10: 'Окт', 11: 'Ноя', 12: 'Дек'}
    chart_months_labels = [f"{ru_months[m.month]} {m.year}" for m in sorted_months]
    chart_months_profits = [float(monthly_data[m]['kirim'] - monthly_data[m]['chiqim']) for m in sorted_months]

    context = {
        'earned_today': earned_today,
        'spent_today': spent_today,
        'profit_today': profit_today,
        'income_month': income_month,
        'income_year': income_year,
        'clients_count': clients_count,
        'purchased_items': purchased_items,
        'cash_balance': cash_balance,
        'top_clients': top_clients,
        'action_history': action_history,
        # Chart JSON configurations
        'chart_days_labels': chart_days_labels,
        'chart_days_profits': chart_days_profits,
        'chart_days_income': chart_days_income,
        'chart_weeks_labels': chart_weeks_labels,
        'chart_weeks_profits': chart_weeks_profits,
        'chart_months_labels': chart_months_labels,
        'chart_months_profits': chart_months_profits,
        'total_kirim_sum': float(total_kirim),
        'total_chiqim_sum': float(total_chiqim),
    }
    return render(request, 'crm/dashboard.html', context)

@login_required
def kirim_view(request):
    if request.method == 'POST':
        client_name = request.POST.get('client_name')
        amount = request.POST.get('amount')
        comment = request.POST.get('comment', '')
        
        if client_name and amount:
            try:
                amount_dec = Decimal(amount)
                Kirim.objects.create(
                    client_name=client_name,
                    amount=amount_dec,
                    comment=comment,
                    author=request.user
                )
                messages.success(request, "Запись прихода успешно добавлена!")
            except Exception as e:
                messages.error(request, f"Ошибка при сохранении: {str(e)}")
        else:
            messages.error(request, "Пожалуйста, заполните обязательные поля.")
        return redirect('kirim')

    # GET request
    if request.user.is_admin():
        kirim_list = Kirim.objects.all().order_by('-created_at')
    else:
        kirim_list = Kirim.objects.filter(author=request.user).order_by('-created_at')
        
    paginator = Paginator(kirim_list, 15)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, 'crm/kirim.html', {'page_obj': page_obj})

@login_required
def chiqim_view(request):
    if request.method == 'POST':
        item_name = request.POST.get('item_name')
        quantity = request.POST.get('quantity')
        price = request.POST.get('price')
        comment = request.POST.get('comment', '')
        
        if item_name and quantity and price:
            try:
                qty_val = int(quantity)
                price_val = Decimal(price)
                Chiqim.objects.create(
                    item_name=item_name,
                    quantity=qty_val,
                    price=price_val,
                    comment=comment,
                    author=request.user
                )
                messages.success(request, "Запись расхода успешно добавлена!")
            except Exception as e:
                messages.error(request, f"Ошибка при сохранении: {str(e)}")
        else:
            messages.error(request, "Пожалуйста, заполните обязательные поля.")
        return redirect('chiqim')

    # GET request
    if request.user.is_admin():
        chiqim_list = Chiqim.objects.all().order_by('-created_at')
    else:
        chiqim_list = Chiqim.objects.filter(author=request.user).order_by('-created_at')
        
    paginator = Paginator(chiqim_list, 15)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, 'crm/chiqim.html', {'page_obj': page_obj})

@login_required
def search_view(request):
    q = request.GET.get('q', '')
    
    # Filter base objects depending on role
    if request.user.is_admin():
        kirim_qs = Kirim.objects.all()
        chiqim_qs = Chiqim.objects.all()
    else:
        kirim_qs = Kirim.objects.filter(author=request.user)
        chiqim_qs = Chiqim.objects.filter(author=request.user)
        
    kirim_results = []
    chiqim_results = []
    
    if q:
        # Check if query is a valid date (YYYY-MM-DD or similar)
        # Or simple partial searches
        # Search Kirim: client name, author, date, comment
        kirim_results = kirim_qs.filter(
            Q(client_name__icontains=q) |
            Q(author__username__icontains=q) |
            Q(comment__icontains=q) |
            Q(created_at__icontains=q)
        ).order_by('-created_at')
        
        # Search Chiqim: item name, author, date, comment
        chiqim_results = chiqim_qs.filter(
            Q(item_name__icontains=q) |
            Q(author__username__icontains=q) |
            Q(comment__icontains=q) |
            Q(created_at__icontains=q)
        ).order_by('-created_at')
        
    return render(request, 'crm/search.html', {
        'q': q,
        'kirim_results': kirim_results,
        'chiqim_results': chiqim_results
    })

@login_required
@admin_required
def reports_view(request):
    period = request.GET.get('period', 'today')
    start_date, end_date = get_date_range(period)
    
    kirim_qs = Kirim.objects.all()
    chiqim_qs = Chiqim.objects.all()
    
    if start_date and end_date:
        kirim_qs = kirim_qs.filter(created_at__range=(start_date, end_date))
        chiqim_qs = chiqim_qs.filter(created_at__range=(start_date, end_date))
        
    total_kirim = kirim_qs.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    total_chiqim = chiqim_qs.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    profit = total_kirim - total_chiqim

    # Cash balance is ALWAYS the running overall balance of the clinic
    all_kirim = Kirim.objects.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    all_chiqim = Chiqim.objects.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    cash_balance = all_kirim - all_chiqim

    # Details table
    recent_kirims = kirim_qs.order_by('-created_at')[:10]
    recent_chiqims = chiqim_qs.order_by('-created_at')[:10]

    return render(request, 'crm/reports.html', {
        'period': period,
        'total_kirim': total_kirim,
        'total_chiqim': total_chiqim,
        'profit': profit,
        'cash_balance': cash_balance,
        'recent_kirims': recent_kirims,
        'recent_chiqims': recent_chiqims
    })

@login_required
@admin_required
def export_excel(request):
    period = request.GET.get('period', 'today')
    start_date, end_date = get_date_range(period)
    
    kirim_qs = Kirim.objects.all()
    chiqim_qs = Chiqim.objects.all()
    
    if start_date and end_date:
        kirim_qs = kirim_qs.filter(created_at__range=(start_date, end_date))
        chiqim_qs = chiqim_qs.filter(created_at__range=(start_date, end_date))
        
    total_kirim = kirim_qs.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    total_chiqim = chiqim_qs.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    profit = total_kirim - total_chiqim

    wb = Workbook()
    
    # Sheet 1: Summary (Сводка)
    ws_summary = wb.active
    ws_summary.title = "Сводка"
    ws_summary.views.sheetView[0].showGridLines = True
    
    # Styles
    title_font = Font(name='Arial', size=16, bold=True, color='0D6EFD')
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    bold_font = Font(name='Arial', size=11, bold=True)
    normal_font = Font(name='Arial', size=11)
    
    header_fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')
    sum_fill = PatternFill(start_color='E8F0FE', end_color='E8F0FE', fill_type='solid')
    border_side = Side(border_style="thin", color="CCCCCC")
    border_double = Side(border_style="double", color="333333")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    double_bottom_border = Border(top=border_side, bottom=border_double)
    
    # Summary title
    ws_summary['A1'] = f"Финансовый отчет клиники за период: {period.upper()}"
    ws_summary['A1'].font = title_font
    ws_summary.row_dimensions[1].height = 30
    
    # Summary Table Headers
    ws_summary['A3'] = "Показатель"
    ws_summary['B3'] = "Значение (сум)"
    ws_summary['A3'].font = header_font
    ws_summary['A3'].fill = header_fill
    ws_summary['B3'].font = header_font
    ws_summary['B3'].fill = header_fill
    ws_summary['A3'].alignment = Alignment(horizontal='left')
    ws_summary['B3'].alignment = Alignment(horizontal='right')
    
    summary_data = [
        ("Общий Приход (Kirim)", float(total_kirim)),
        ("Общий Расход (Chiqim)", float(total_chiqim)),
        ("Общая Прибыль", float(profit)),
    ]
    
    for row_idx, (name, val) in enumerate(summary_data, start=4):
        ws_summary.cell(row=row_idx, column=1, value=name).font = normal_font
        ws_summary.cell(row=row_idx, column=1).border = thin_border
        
        val_cell = ws_summary.cell(row=row_idx, column=2, value=val)
        val_cell.font = bold_font if "Прибыль" in name else normal_font
        val_cell.number_format = '#,##0.00'
        val_cell.border = thin_border
        val_cell.alignment = Alignment(horizontal='right')
        
    # Auto-adjust column widths
    for col in ws_summary.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = max(max_len + 3, 20)
        
    # Sheet 2: Kirim
    ws_kirim = wb.create_sheet(title="Приход (Kirim)")
    ws_kirim.views.sheetView[0].showGridLines = True
    ws_kirim.append(["Имя клиента", "Сумма (сум)", "Комментарий", "Дата и время", "Кто добавил"])
    for cell in ws_kirim[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        
    for k in kirim_qs:
        local_time = timezone.localtime(k.created_at).strftime('%d.%m.%Y %H:%M')
        row = [k.client_name, float(k.amount), k.comment or "", local_time, k.author.username]
        ws_kirim.append(row)
        
    for r in range(2, ws_kirim.max_row + 1):
        ws_kirim.cell(row=r, column=2).number_format = '#,##0.00'
        ws_kirim.cell(row=r, column=2).alignment = Alignment(horizontal='right')
        for c in range(1, 6):
            ws_kirim.cell(row=r, column=c).border = thin_border
            ws_kirim.cell(row=r, column=c).font = normal_font
            
    # Auto width for Sheet 2
    for col in ws_kirim.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_kirim.column_dimensions[col_letter].width = max(max_len + 3, 15)

    # Sheet 3: Chiqim
    ws_chiqim = wb.create_sheet(title="Расход (Chiqim)")
    ws_chiqim.views.sheetView[0].showGridLines = True
    ws_chiqim.append(["Название товара", "Количество", "Цена (сум)", "Общая стоимость (сум)", "Комментарий", "Дата и время", "Кто добавил"])
    for cell in ws_chiqim[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        
    for c in chiqim_qs:
        local_time = timezone.localtime(c.created_at).strftime('%d.%m.%Y %H:%M')
        row = [c.item_name, c.quantity, float(c.price), float(c.total_price), c.comment or "", local_time, c.author.username]
        ws_chiqim.append(row)
        
    for r in range(2, ws_chiqim.max_row + 1):
        ws_chiqim.cell(row=r, column=2).alignment = Alignment(horizontal='center')
        ws_chiqim.cell(row=r, column=3).number_format = '#,##0.00'
        ws_chiqim.cell(row=r, column=3).alignment = Alignment(horizontal='right')
        ws_chiqim.cell(row=r, column=4).number_format = '#,##0.00'
        ws_chiqim.cell(row=r, column=4).alignment = Alignment(horizontal='right')
        for col_idx in range(1, 8):
            ws_chiqim.cell(row=r, column=col_idx).border = thin_border
            ws_chiqim.cell(row=r, column=col_idx).font = normal_font

    # Auto width for Sheet 3
    for col in ws_chiqim.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_chiqim.column_dimensions[col_letter].width = max(max_len + 3, 15)

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename=report_{period}_{timezone.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    wb.save(response)
    return response

@login_required
@admin_required
def export_pdf(request):
    period = request.GET.get('period', 'today')
    start_date, end_date = get_date_range(period)
    
    kirim_qs = Kirim.objects.all()
    chiqim_qs = Chiqim.objects.all()
    
    if start_date and end_date:
        kirim_qs = kirim_qs.filter(created_at__range=(start_date, end_date))
        chiqim_qs = chiqim_qs.filter(created_at__range=(start_date, end_date))
        
    total_kirim = kirim_qs.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    total_chiqim = chiqim_qs.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    profit = total_kirim - total_chiqim

    # Overall cash register balance
    all_kirim = Kirim.objects.aggregate(sum=Sum('amount'))['sum'] or Decimal(0)
    all_chiqim = Chiqim.objects.aggregate(sum=Sum('total_price'))['sum'] or Decimal(0)
    cash_balance = all_kirim - all_chiqim

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename=report_{period}_{timezone.now().strftime("%Y%m%d_%H%M%S")}.pdf'

    doc = SimpleDocTemplate(response, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []

    # Setup styles
    styles = getSampleStyleSheet()
    
    # Choose registered Arial font if Cyrillic supported, otherwise default standard Helvetica
    font_name = 'Arial' if CYRILLIC_FONT_SUPPORTED else 'Helvetica'
    font_bold = 'Arial-Bold' if CYRILLIC_FONT_SUPPORTED else 'Helvetica-Bold'

    title_style = ParagraphStyle(
        name='TitleStyle',
        fontName=font_bold,
        fontSize=18,
        textColor=colors.HexColor('#0D6EFD'),
        alignment=1, # Center
        spaceAfter=20
    )
    
    heading_style = ParagraphStyle(
        name='HeadingStyle',
        fontName=font_bold,
        fontSize=14,
        textColor=colors.HexColor('#212529'),
        spaceBefore=15,
        spaceAfter=10
    )

    body_style = ParagraphStyle(
        name='BodyStyle',
        fontName=font_name,
        fontSize=10,
        textColor=colors.HexColor('#495057')
    )

    # Document Header
    story.append(Paragraph("ФИНАНСОВЫЙ ОТЧЕТ ЛОР-КЛИНИКИ", title_style))
    story.append(Paragraph(f"<b>Период:</b> {period.capitalize()}", body_style))
    story.append(Paragraph(f"<b>Дата экспорта:</b> {timezone.now().strftime('%d.%m.%Y %H:%M')}", body_style))
    story.append(Spacer(1, 15))

    # Summary Table Data
    summary_data = [
        [Paragraph(f"<b>Показатель</b>", body_style), Paragraph(f"<b>Сумма (сум)</b>", body_style)],
        [Paragraph("Общий Приход (Kirim)", body_style), Paragraph(f"{total_kirim:,.2f}", body_style)],
        [Paragraph("Общий Расход (Chiqim)", body_style), Paragraph(f"{total_chiqim:,.2f}", body_style)],
        [Paragraph("Чистая прибыль", body_style), Paragraph(f"<b>{profit:,.2f}</b>", body_style)],
        [Paragraph("Остаток в кассе (общий)", body_style), Paragraph(f"<b>{cash_balance:,.2f}</b>", body_style)]
    ]

    t_summary = Table(summary_data, colWidths=[250, 200])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F8F9FA')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DEE2E6')),
        ('BACKGROUND', (0,-2), (-1,-1), colors.HexColor('#E8F0FE')),
    ]))
    
    story.append(t_summary)
    story.append(Spacer(1, 25))

    # Kirim Details
    story.append(Paragraph("Детализация приходов (Kirim) - последние 10 записей", heading_style))
    kirim_header = [
        Paragraph(f"<b>Имя клиента</b>", body_style),
        Paragraph(f"<b>Сумма (сум)</b>", body_style),
        Paragraph(f"<b>Дата</b>", body_style),
        Paragraph(f"<b>Автор</b>", body_style)
    ]
    kirim_table_data = [kirim_header]
    for k in kirim_qs.order_by('-created_at')[:10]:
        local_time = timezone.localtime(k.created_at).strftime('%d.%m.%Y %H:%M')
        kirim_table_data.append([
            Paragraph(k.client_name, body_style),
            Paragraph(f"{k.amount:,.2f}", body_style),
            Paragraph(local_time, body_style),
            Paragraph(k.author.username, body_style)
        ])
        
    t_kirim = Table(kirim_table_data, colWidths=[150, 100, 100, 100])
    t_kirim.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E9F7EF')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DEE2E6')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_kirim)
    story.append(Spacer(1, 20))

    # Chiqim Details
    story.append(Paragraph("Детализация расходов (Chiqim) - последние 10 записей", heading_style))
    chiqim_header = [
        Paragraph(f"<b>Товар</b>", body_style),
        Paragraph(f"<b>Кол-во</b>", body_style),
        Paragraph(f"<b>Цена</b>", body_style),
        Paragraph(f"<b>Итого</b>", body_style),
        Paragraph(f"<b>Дата</b>", body_style),
        Paragraph(f"<b>Автор</b>", body_style)
    ]
    chiqim_table_data = [chiqim_header]
    for c in chiqim_qs.order_by('-created_at')[:10]:
        local_time = timezone.localtime(c.created_at).strftime('%d.%m.%Y %H:%M')
        chiqim_table_data.append([
            Paragraph(c.item_name, body_style),
            Paragraph(str(c.quantity), body_style),
            Paragraph(f"{c.price:,.2f}", body_style),
            Paragraph(f"{c.total_price:,.2f}", body_style),
            Paragraph(local_time, body_style),
            Paragraph(c.author.username, body_style)
        ])
        
    t_chiqim = Table(chiqim_table_data, colWidths=[110, 50, 75, 75, 75, 65])
    t_chiqim.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FDEDEC')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DEE2E6')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_chiqim)

    doc.build(story)
    return response

@login_required
@admin_required
def workers_view(request):
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'create':
            username = request.POST.get('username')
            email = request.POST.get('email')
            passw = request.POST.get('password')
            role = request.POST.get('role', 'worker')
            
            if username and passw:
                if User.objects.filter(username=username).exists():
                    messages.error(request, "Пользователь с таким именем уже существует.")
                else:
                    User.objects.create_user(
                        username=username,
                        email=email,
                        password=passw,
                        role=role
                    )
                    messages.success(request, f"Пользователь {username} успешно создан!")
            else:
                messages.error(request, "Необходимо указать имя пользователя и пароль.")
                
        elif action == 'toggle_status':
            user_id = request.POST.get('user_id')
            user_to_mod = get_object_or_404(User, id=user_id)
            if user_to_mod != request.user: # Cannot deactivate yourself
                user_to_mod.is_active = not user_to_mod.is_active
                user_to_mod.save()
                messages.success(request, f"Статус пользователя {user_to_mod.username} изменен.")
            else:
                messages.error(request, "Вы не можете деактивировать самого себя.")
                
        return redirect('workers')

    workers = User.objects.all().order_by('-date_joined')
    return render(request, 'crm/workers.html', {'workers': workers})

@login_required
@admin_required
def history_view(request):
    # Retrieve all Kirim and Chiqim, sort them chronologically
    kirims = Kirim.objects.all()
    chiqims = Chiqim.objects.all()
    
    actions = []
    for k in kirims:
        actions.append({
            'type': 'kirim',
            'author': k.author,
            'client_or_item': k.client_name,
            'amount': k.amount,
            'comment': k.comment,
            'created_at': k.created_at
        })
    for c in chiqims:
        actions.append({
            'type': 'chiqim',
            'author': c.author,
            'client_or_item': f"{c.item_name} ({c.quantity} шт. х {c.price} сум)",
            'amount': c.total_price,
            'comment': c.comment,
            'created_at': c.created_at
        })
        
    actions = sorted(actions, key=lambda x: x['created_at'], reverse=True)
    
    paginator = Paginator(actions, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    return render(request, 'crm/history.html', {'page_obj': page_obj})

@login_required
@admin_required
def settings_view(request):
    if request.method == 'POST':
        # Handles password change for the current user
        form = PasswordChangeForm(request.user, request.POST)
        if form.is_valid():
            user = form.save()
            update_session_auth_hash(request, user)  # Keep the user logged in
            messages.success(request, 'Пароль успешно обновлен!')
            return redirect('settings')
        else:
            messages.error(request, 'Пожалуйста, исправьте ошибки ниже.')
    else:
        form = PasswordChangeForm(request.user)
        
    return render(request, 'crm/settings.html', {'form': form})

