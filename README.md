import openpyxl

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand
from app.models import *


class Command(BaseCommand):
    help = 'Импорт из Exel'

    def handle(self, *args, **options):
        User = get_user_model()
        wb = openpyxl.load_workbook('import/user_import.xlsx')
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            role = row[0]
            fio = row[1]
            login = row[2]
            password = row[3]

            user, created = User.objects.get_or_create(username=login)

            user.email = login
            user.first_name = ""
            user.last_name = fio
            user.is_staff = role in ["Администратор", "Менеджер"]
            user.is_superuser = role == "Администратор"
            user.is_active = True

            if created:
                user.set_password(str(password))

            user.save()

        self.stdout.write('Импорт завершён')

        wb2 = openpyxl.load_workbook('import/Tovar.xlsx')
        ws2 = wb2.active
        
        for row in ws2.iter_rows(min_row=2, values_only=True):
            edinizaIzmernia, _ = EdinizaIzmernia.objects.get_or_create(nazvanie=row[2])
            postav, _ = Postavshik.objects.get_or_create(nazvanie=row[4])
            proizvoditel, _ = Proizvoditel.objects.get_or_create(nazvanie=row[5])
            kategoria, _ = Kategoria.objects.get_or_create(nazvanie=row[6])
            
            Tovar.objects.create(
                artikul = row[0],
                nazvanie = row[1],
                edinizaIzmernia = edinizaIzmernia,
                cena = int(row[3]),
                postavshik = postav,
                proizvoditel = proizvoditel,
                kategoria = kategoria,
                skidka = int(row[7]),
                kolichestvo = int(row[8]),
                opisanie = row[9],
                foto = row[9],
            )


# `views.py`

```python
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.db.models import Q

from .models import Product
from .forms import ProductForm


def guest_entry(request):
    return redirect("product_list")

```

# `product_list`


```python
def product_list(request):
    # Берём все товары из базы
    products = Product.objects.all()


    # 1. ПОЛУЧАЕМ ДАННЫЕ ИЗ GET

    # Поиск
    # HTML: <input name="search">
    search = request.GET.get("search", "").strip()

    # Фильтр по поставщику
    # HTML: <select name="postavshik">
    # "all" значит "все поставщики", то есть фильтр не применяем
    postavshik = request.GET.get("postavshik", "all")

    # Фильтр по производителю
    proizvoditel = request.GET.get("proizvoditel", "all")

    # Фильтр по категории
    kategoriya = request.GET.get("kategoriya", "all")

    # Скидка от/до
    # HTML: <input name="skidka_min">
    # HTML: <input name="skidka_max">
    # skidka_min=10 значит скидка от 10
    # skidka_max=30 значит скидка до 30
    skidka_min = request.GET.get("skidka_min", "").strip()
    skidka_max = request.GET.get("skidka_max", "").strip()

    # Цена от/до
    cena_min = request.GET.get("cena_min", "").strip()
    cena_max = request.GET.get("cena_max", "").strip()

    # Количество на складе от/до
    kolichestvo_min = request.GET.get("kolichestvo_min", "").strip()
    kolichestvo_max = request.GET.get("kolichestvo_max", "").strip()

    # Чекбокс "только в наличии"
    # Если галочка стоит, придёт значение "yes"
    v_nalichii = request.GET.get("v_nalichii", "")

    # Чекбокс "только со скидкой"
    so_skidkoi = request.GET.get("so_skidkoi", "")

    # Сортировка
    # Например:
    # sort=cena    -> цена по возрастанию
    # sort=-cena   -> цена по убыванию
    sort = request.GET.get("sort", "")

    # Списки для выпадающих фильтров в HTML
    postavshiki = []
    proizvoditeli = []
    kategorii = []

    # 2. ФИЛЬТРЫ И СОРТИРОВКА ТОЛЬКО МЕНЕДЖЕРУ И АДМИНУ

    if request.user.is_staff or request.user.is_superuser:

        # ПОИСК ПО ТЕКСТОВЫМ ПОЛЯМ
        if search:
            products = products.filter(
                Q(artikul__icontains=search) |
                Q(nazvanie__icontains=search) |
                Q(edinica__icontains=search) |
                Q(postavshik__icontains=search) |
                Q(proizvoditel__icontains=search) |
                Q(kategoriya__icontains=search) |
                Q(opisanie__icontains=search)
            )

        # ФИЛЬТР ПО ПОСТАВЩИКУ
        if postavshik != "all":
            products = products.filter(postavshik=postavshik)


        # ФИЛЬТР ПО ПРОИЗВОДИТЕЛЮ
        if proizvoditel != "all":
            products = products.filter(proizvoditel=proizvoditel)

        # ФИЛЬТР ПО КАТЕГОРИИ
        if kategoriya != "all":
            products = products.filter(kategoriya=kategoriya)

        # ФИЛЬТР ПО СКИДКЕ
        # __gte = больше или равно
        # skidka__gte=10 значит скидка >= 10
        if skidka_min:
            products = products.filter(skidka__gte=skidka_min)

        # __lte = меньше или равно
        # skidka__lte=30 значит скидка <= 30
        if skidka_max:
            products = products.filter(skidka__lte=skidka_max)

        # ФИЛЬТР ПО ЦЕНЕ
        if cena_min:
            products = products.filter(cena__gte=cena_min)

        if cena_max:
            products = products.filter(cena__lte=cena_max)

        # ФИЛЬТР ПО КОЛИЧЕСТВУ
        if kolichestvo_min:
            products = products.filter(kolichestvo__gte=kolichestvo_min)

        if kolichestvo_max:
            products = products.filter(kolichestvo__lte=kolichestvo_max)

        # -------------------------
        # ТОЛЬКО В НАЛИЧИИ
        # -------------------------
        # __gt = больше
        # kolichestvo__gt=0 значит количество > 0
        if v_nalichii == "yes":
            products = products.filter(kolichestvo__gt=0)

        # -------------------------
        # ТОЛЬКО СО СКИДКОЙ
        # -------------------------
        if so_skidkoi == "yes":
            products = products.filter(skidka__gt=0)

        # -------------------------
        # СОРТИРОВКА
        # -------------------------
        allowed_sorts = [
            "nazvanie",
            "-nazvanie",
            "cena",
            "-cena",
            "kolichestvo",
            "-kolichestvo",
            "skidka",
            "-skidka",
            "artikul",
            "-artikul",
        ]

        if sort in allowed_sorts:
            products = products.order_by(sort)

        # -------------------------
        # ДАННЫЕ ДЛЯ SELECT В HTML
        # -------------------------
        postavshiki = Product.objects.values_list("postavshik", flat=True).distinct()
        proizvoditeli = Product.objects.values_list("proizvoditel", flat=True).distinct()
        kategorii = Product.objects.values_list("kategoriya", flat=True).distinct()

    # =========================
    # 3. ОТПРАВЛЯЕМ ДАННЫЕ В HTML
    # =========================

    return render(request, "product_list.html", {
        "products": products,

        "search": search,

        "postavshik": postavshik,
        "proizvoditel": proizvoditel,
        "kategoriya": kategoriya,

        "skidka_min": skidka_min,
        "skidka_max": skidka_max,

        "cena_min": cena_min,
        "cena_max": cena_max,

        "kolichestvo_min": kolichestvo_min,
        "kolichestvo_max": kolichestvo_max,

        "v_nalichii": v_nalichii,
        "so_skidkoi": so_skidkoi,

        "sort": sort,

        "postavshiki": postavshiki,
        "proizvoditeli": proizvoditeli,
        "kategorii": kategorii,
    })
```

---

# Минимальный `product_list`

```text
гость/клиент видят товары
менеджер/админ могут искать, фильтровать по поставщику и сортировать
админ может CRUD
```

```python
def product_list(request):
    products = Product.objects.all()

    search = request.GET.get("search", "").strip()
    postavshik = request.GET.get("postavshik", "all")
    sort = request.GET.get("sort", "")

    postavshiki = []

    if request.user.is_staff or request.user.is_superuser:
        if search:
            products = products.filter(
                Q(artikul__icontains=search) |
                Q(nazvanie__icontains=search) |
                Q(opisanie__icontains=search) |
                Q(postavshik__icontains=search) |
                Q(proizvoditel__icontains=search) |
                Q(kategoriya__icontains=search)
            )

        if postavshik != "all":
            products = products.filter(postavshik=postavshik)

        allowed_sorts = [
            "nazvanie",
            "-nazvanie",
            "cena",
            "-cena",
            "kolichestvo",
            "-kolichestvo",
            "skidka",
            "-skidka",
        ]

        if sort in allowed_sorts:
            products = products.order_by(sort)

        postavshiki = Product.objects.values_list("postavshik", flat=True).distinct()

    return render(request, "product_list.html", {
        "products": products,
        "search": search,
        "postavshik": postavshik,
        "sort": sort,
        "postavshiki": postavshiki,
    })
```


# CRUD

## Добавление

```python
@login_required
def product_create(request):
    # Добавлять может только админ
    if not request.user.is_superuser:
        return redirect("product_list")

    # Если нажали кнопку "Сохранить"
    if request.method == "POST":
        form = ProductForm(request.POST, request.FILES)

        if form.is_valid():
            form.save()
            return redirect("product_list")

    # Если просто открыли страницу
    else:
        form = ProductForm()

    return render(request, "product_form.html", {
        "form": form,
        "title": "Добавить товар",
    })
```

---

## Редактирование товара

```python
@login_required
def product_update(request, pk):
    # Редактировать может только админ
    if not request.user.is_superuser:
        return redirect("product_list")

    # Ищем товар по id
    product = get_object_or_404(Product, pk=pk)

    # Если нажали "Сохранить"
    if request.method == "POST":
        # instance=product значит редактируем существующий товар,
        # а не создаём новый
        form = ProductForm(request.POST, request.FILES, instance=product)

        if form.is_valid():
            form.save()
            return redirect("product_list")

    # Если просто открыли страницу
    else:
        form = ProductForm(instance=product)

    return render(request, "product_form.html", {
        "form": form,
        "title": "Редактировать товар",
    })
```

---

## Удаление товара

```python
@login_required
def product_delete(request, pk):
    # Удалять может только админ
    if not request.user.is_superuser:
        return redirect("product_list")

    # Ищем товар по id
    product = get_object_or_404(Product, pk=pk)

    # Если нажали кнопку "Да, удалить"
    if request.method == "POST":
        product.delete()
        return redirect("product_list")

    # Если просто открыли страницу удаления
    return render(request, "confirm_delete.html", {
        "product": product,
    })
```


---

# 7. `product_list.html`

`templates/product_list.html`

```html
{% extends "base.html" %}
{% load static %}

{% block title %}Список товаров{% endblock %}

{% block content %}
<h1>Список товаров</h1>

<!--
Фильтры показываем только менеджеру и админу.
-->
{% if user.is_staff or user.is_superuser %}
<form method="get">

    <!-- ПОИСК -->
    <p>
        <input type="text" name="search" value="{{ search }}" placeholder="Поиск">
    </p>

    <!-- ФИЛЬТРЫ ПО СПИСКАМ -->
    <p>
        <select name="postavshik">
            <option value="all">Все поставщики</option>
            {% for p in postavshiki %}
                <option value="{{ p }}" {% if postavshik == p %}selected{% endif %}>
                    {{ p }}
                </option>
            {% endfor %}
        </select>

        <select name="proizvoditel">
            <option value="all">Все производители</option>
            {% for p in proizvoditeli %}
                <option value="{{ p }}" {% if proizvoditel == p %}selected{% endif %}>
                    {{ p }}
                </option>
            {% endfor %}
        </select>

        <select name="kategoriya">
            <option value="all">Все категории</option>
            {% for k in kategorii %}
                <option value="{{ k }}" {% if kategoriya == k %}selected{% endif %}>
                    {{ k }}
                </option>
            {% endfor %}
        </select>
    </p>

    <!-- СКИДКА ОТ/ДО -->
    <p>
        Скидка от:
        <input type="number" name="skidka_min" value="{{ skidka_min }}">
        до:
        <input type="number" name="skidka_max" value="{{ skidka_max }}">
    </p>

    <!-- ЦЕНА ОТ/ДО -->
    <p>
        Цена от:
        <input type="number" name="cena_min" value="{{ cena_min }}">
        до:
        <input type="number" name="cena_max" value="{{ cena_max }}">
    </p>

    <!-- КОЛИЧЕСТВО ОТ/ДО -->
    <p>
        Количество от:
        <input type="number" name="kolichestvo_min" value="{{ kolichestvo_min }}">
        до:
        <input type="number" name="kolichestvo_max" value="{{ kolichestvo_max }}">
    </p>

    <!-- ЧЕКБОКСЫ -->
    <p>
        <label>
            <input type="checkbox" name="v_nalichii" value="yes" {% if v_nalichii == "yes" %}checked{% endif %}>
            Только в наличии
        </label>

        <label>
            <input type="checkbox" name="so_skidkoi" value="yes" {% if so_skidkoi == "yes" %}checked{% endif %}>
            Только со скидкой
        </label>
    </p>

    <!-- СОРТИРОВКА -->
    <p>
        <select name="sort">
            <option value="">Без сортировки</option>

            <option value="nazvanie" {% if sort == "nazvanie" %}selected{% endif %}>
                Название ↑
            </option>

            <option value="-nazvanie" {% if sort == "-nazvanie" %}selected{% endif %}>
                Название ↓
            </option>

            <option value="cena" {% if sort == "cena" %}selected{% endif %}>
                Цена ↑
            </option>

            <option value="-cena" {% if sort == "-cena" %}selected{% endif %}>
                Цена ↓
            </option>

            <option value="kolichestvo" {% if sort == "kolichestvo" %}selected{% endif %}>
                Количество ↑
            </option>

            <option value="-kolichestvo" {% if sort == "-kolichestvo" %}selected{% endif %}>
                Количество ↓
            </option>

            <option value="skidka" {% if sort == "skidka" %}selected{% endif %}>
                Скидка ↑
            </option>

            <option value="-skidka" {% if sort == "-skidka" %}selected{% endif %}>
                Скидка ↓
            </option>

            <option value="artikul" {% if sort == "artikul" %}selected{% endif %}>
                Артикул ↑
            </option>

            <option value="-artikul" {% if sort == "-artikul" %}selected{% endif %}>
                Артикул ↓
            </option>
        </select>

        <button class="btn" type="submit">Применить</button>
        <a class="btn" href="{% url 'product_list' %}">Сбросить</a>
    </p>
</form>
{% endif %}

<!-- Кнопка добавления только админу -->
{% if user.is_superuser %}
    <p>
        <a class="btn" href="{% url 'product_create' %}">Добавить товар</a>
    </p>
{% endif %}

<!-- СПИСОК ТОВАРОВ -->
{% for product in products %}

<!--
Если товара нет на складе — голубой фон.
Иначе если скидка больше 15 — зелёный фон.
-->
<div class="card {% if product.kolichestvo == 0 %}no-stock{% elif product.skidka > 15 %}sale{% endif %}">

    <div>
        {% if product.foto %}
            <img class="product-img" src="{{ product.foto.url }}" alt="{{ product.nazvanie }}">
        {% else %}
            <img class="product-img" src="{% static 'picture.png' %}" alt="Заглушка">
        {% endif %}
    </div>

    <div>
        <h2>{{ product.nazvanie }}</h2>

        <p><b>Артикул:</b> {{ product.artikul }}</p>
        <p><b>Категория:</b> {{ product.kategoriya }}</p>
        <p><b>Описание:</b> {{ product.opisanie }}</p>
        <p><b>Производитель:</b> {{ product.proizvoditel }}</p>
        <p><b>Поставщик:</b> {{ product.postavshik }}</p>
        <p><b>Единица измерения:</b> {{ product.edinica }}</p>
        <p><b>Количество на складе:</b> {{ product.kolichestvo }}</p>
        <p><b>Скидка:</b> {{ product.skidka }}%</p>

        {% if product.skidka > 0 %}
            <p>
                <b>Цена:</b>
                <span class="old-price">{{ product.cena }}</span>
                <span class="new-price">{{ product.cena_so_skidkoi|floatformat:2 }}</span>
            </p>
        {% else %}
            <p><b>Цена:</b> {{ product.cena }}</p>
        {% endif %}

        <!-- Кнопки CRUD только админу -->
        {% if user.is_superuser %}
            <a class="btn" href="{% url 'product_update' product.pk %}">Редактировать</a>
            <a class="btn" href="{% url 'product_delete' product.pk %}">Удалить</a>
        {% endif %}
    </div>
</div>

{% empty %}
<p>Товары не найдены</p>
{% endfor %}

{% endblock %}
```


---

# `confirm_delete.html`

```html
{% extends "base.html" %}

{% block title %}Удаление{% endblock %}

{% block content %}
<h1>Удалить товар</h1>

<p>Точно удалить товар: <b>{{ product.nazvanie }}</b>?</p>

<form method="post">
    {% csrf_token %}
    <button class="btn" type="submit">Да, удалить</button>
    <a class="btn" href="{% url 'product_list' %}">Назад</a>
</form>
{% endblock %}
```

---

# `base.html`

```html
{% load static %}
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>{% block title %}Обувь{% endblock %}</title>
    <link rel="icon" href="{% static 'Icon.ico' %}">

    <style>
        body {
            font-family: "Times New Roman", Times, serif;
            background: white;
            margin: 0;
            padding: 0;
        }

        header {
            background: #7FFF00;
            padding: 10px;
        }

        header img {
            height: 70px;
            width: auto;
            vertical-align: middle;
        }

        nav {
            display: inline-block;
            margin-left: 20px;
        }

        nav a {
            margin-right: 15px;
            color: black;
            font-weight: bold;
            text-decoration: none;
        }

        main {
            padding: 20px;
        }

        .btn {
            background: #00FA9A;
            border: 1px solid black;
            padding: 6px 10px;
            color: black;
            text-decoration: none;
            display: inline-block;
            margin: 3px;
        }

        .card {
            border: 1px solid black;
            margin-bottom: 10px;
            padding: 10px;
            display: flex;
            gap: 15px;
        }

        .sale {
            background: #2E8B57;
        }

        .no-stock {
            background: lightblue;
        }

        .old-price {
            color: red;
            text-decoration: line-through;
        }

        .new-price {
            color: black;
            font-weight: bold;
        }

        .product-img {
            width: 120px;
            height: 120px;
            object-fit: contain;
        }

        input, select {
            margin: 3px;
        }
    </style>
</head>
<body>

<header>
    <img src="{% static 'Icon.png' %}" alt="Логотип">

    <nav>
        <a href="{% url 'product_list' %}">Товары</a>

        {% if user.is_superuser %}
            <a href="{% url 'product_create' %}">Добавить товар</a>
        {% endif %}

        {% if user.is_authenticated %}
            <span>{{ user.last_name }} {{ user.first_name }}</span>
            <a href="{% url 'logout' %}">Выйти</a>
        {% else %}
            <a href="{% url 'login' %}">Войти</a>
        {% endif %}
    </nav>
</header>

<main>
    {% block content %}{% endblock %}
</main>

</body>
</html>
```

---


## Если гость и клиент могут только смотреть товары

`product_list` без `login_required`.

Кнопки CRUD в HTML:

```html
{% if user.is_superuser %}
    <a href="{% url 'product_update' product.pk %}">Редактировать</a>
    <a href="{% url 'product_delete' product.pk %}">Удалить</a>
{% endif %}
```

---

## Если поиск/фильтры только менеджеру и админу

В `views.py`:

```python
if request.user.is_staff or request.user.is_superuser:
    # search/filter/sort
```

В HTML:

```html
{% if user.is_staff or user.is_superuser %}
    <form method="get">
        ...
    </form>
{% endif %}
```

---

## Если поиск доступен всем

Во `views.py` убрать поиск из проверки `is_staff`.

Было:

```python
if request.user.is_staff or request.user.is_superuser:
    if search:
        products = products.filter(...)
```

Стало:

```python
if search:
    products = products.filter(...)
```

---

## Если CRUD только админу

Во `views.py`:

```python
if not request.user.is_superuser:
    return redirect("product_list")
```

В HTML:

```html
{% if user.is_superuser %}
    кнопки добавить/редактировать/удалить
{% endif %}
```

---

## Если CRUD менеджеру и админу

Во `views.py`:

```python
if not (request.user.is_staff or request.user.is_superuser):
    return redirect("product_list")
```

В HTML:

```html
{% if user.is_staff or user.is_superuser %}
    кнопки добавить/редактировать/удалить
{% endif %}
```

---

# Конструктор фильтров

## Поиск по одному полю

```python
if search:
    products = products.filter(nazvanie__icontains=search)
```

## Поиск по нескольким полям

```python
if search:
    products = products.filter(
        Q(nazvanie__icontains=search) |
        Q(artikul__icontains=search) |
        Q(opisanie__icontains=search)
    )
```

## Фильтр по точному значению

```python
if postavshik != "all":
    products = products.filter(postavshik=postavshik)
```

## Фильтр “от”

```python
if cena_min:
    products = products.filter(cena__gte=cena_min)
```

## Фильтр “до”

```python
if cena_max:
    products = products.filter(cena__lte=cena_max)
```

## Больше

```python
products = products.filter(kolichestvo__gt=0)
```

## Меньше

```python
products = products.filter(kolichestvo__lt=5)
```

## Больше или равно

```python
products = products.filter(skidka__gte=10)
```

## Меньше или равно

```python
products = products.filter(skidka__lte=30)
```

## Сортировка

```python
sort = request.GET.get("sort", "")

allowed_sorts = [
    "nazvanie",
    "-nazvanie",
    "cena",
    "-cena",
]

if sort in allowed_sorts:
    products = products.order_by(sort)
```

---

# 14. Конструктор HTML-полей для фильтров

## Поиск

```html
<input type="text" name="search" value="{{ search }}" placeholder="Поиск">
```

Во `views.py`:

```python
search = request.GET.get("search", "").strip()
```

---

## Select “Все поставщики”

```html
<select name="postavshik">
    <option value="all">Все поставщики</option>
    {% for p in postavshiki %}
        <option value="{{ p }}" {% if postavshik == p %}selected{% endif %}>
            {{ p }}
        </option>
    {% endfor %}
</select>
```

Во `views.py`:

```python
postavshik = request.GET.get("postavshik", "all")

if postavshik != "all":
    products = products.filter(postavshik=postavshik)

postavshiki = Product.objects.values_list("postavshik", flat=True).distinct()
```

---

## Число от/до

```html
Цена от:
<input type="number" name="cena_min" value="{{ cena_min }}">

до:
<input type="number" name="cena_max" value="{{ cena_max }}">
```

Во `views.py`:

```python
cena_min = request.GET.get("cena_min", "").strip()
cena_max = request.GET.get("cena_max", "").strip()

if cena_min:
    products = products.filter(cena__gte=cena_min)

if cena_max:
    products = products.filter(cena__lte=cena_max)
```

---

## Чекбокс

```html
<label>
    <input type="checkbox" name="v_nalichii" value="yes" {% if v_nalichii == "yes" %}checked{% endif %}>
    Только в наличии
</label>
```

Во `views.py`:

```python
v_nalichii = request.GET.get("v_nalichii", "")

if v_nalichii == "yes":
    products = products.filter(kolichestvo__gt=0)
```

---

## Сортировка

```html
<select name="sort">
    <option value="">Без сортировки</option>
    <option value="cena" {% if sort == "cena" %}selected{% endif %}>Цена ↑</option>
    <option value="-cena" {% if sort == "-cena" %}selected{% endif %}>Цена ↓</option>
</select>
```

Во `views.py`:

```python
sort = request.GET.get("sort", "")

allowed_sorts = ["cena", "-cena"]

if sort in allowed_sorts:
    products = products.order_by(sort)
```

---
