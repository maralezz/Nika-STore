"""
🔵 NIKA Сантехника — интернет-магазин
Flask + SQLite + Tailwind (синий премиум-дизайн)
"""
import os
import sqlite3
from functools import wraps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, g, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = 'nika-aquastyle-secret-key-2026'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, 'nika.db')


# ================== БАЗА ==================
def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
    return db


@app.teardown_appcontext
def close_db(exc):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    name TEXT DEFAULT '',
    phone TEXT DEFAULT '',
    address TEXT DEFAULT '',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    category TEXT NOT NULL,
    brand TEXT NOT NULL,
    price REAL NOT NULL,
    old_price REAL,
    description TEXT DEFAULT '',
    icon TEXT DEFAULT '🔧',
    color TEXT DEFAULT 'from-blue-500 to-cyan-500',
    material TEXT DEFAULT '',
    diameter TEXT DEFAULT '',
    in_stock INTEGER DEFAULT 1,
    is_new INTEGER DEFAULT 0,
    is_hit INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    total REAL NOT NULL,
    status TEXT DEFAULT 'Новый',
    address TEXT DEFAULT '',
    phone TEXT DEFAULT '',
    comment TEXT DEFAULT '',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users (id)
);
CREATE TABLE IF NOT EXISTS order_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL,
    product_id INTEGER,
    product_name TEXT,
    quantity INTEGER NOT NULL,
    price REAL NOT NULL,
    FOREIGN KEY (order_id) REFERENCES orders (id)
);
CREATE TABLE IF NOT EXISTS favorites (
    user_id INTEGER NOT NULL,
    product_id INTEGER NOT NULL,
    PRIMARY KEY (user_id, product_id)
);
"""

PRODUCTS = [
    ('Смеситель для кухни Grohe Minta', 'Смесители', 'Grohe', 24990, 28990,
     'Хромированный смеситель с выдвижным душем и поворотным изливом. Керамический картридж, экономия воды.',
     '🚰', 'from-blue-500 to-cyan-500', 'Латунь', '1/2"', 1, 0, 1),
    ('Смеситель для ванны Hansgrohe Logis', 'Смесители', 'Hansgrohe', 18500, None,
     'Однорычажный смеситель с аэратором. Гладкая поверхность Chrome.',
     '🚿', 'from-sky-500 to-blue-600', 'Латунь', '1/2"', 1, 0, 1),
    ('Смеситель для раковины Grohe Eurosmart', 'Смесители', 'Grohe', 9900, 11900,
     'Классический однорычажный смеситель. Технология StarLight.',
     '🔧', 'from-indigo-500 to-blue-500', 'Латунь', '3/8"', 1, 0, 0),
    ('Душевая система Grohe Rainshower', 'Душевые системы', 'Grohe', 45900, 52900,
     'Тропический душ 310 мм, ручной душ, термостат. Нержавейка.',
     '🌧️', 'from-cyan-500 to-teal-500', 'Нержавеющая сталь', '1/2"', 1, 1, 1),
    ('Душевой гарнитур Hansgrohe Croma', 'Душевые системы', 'Hansgrohe', 7900, None,
     'Ручной душ с 3 режимами и шланг 1.6 м.',
     '💧', 'from-sky-400 to-cyan-500', 'Пластик/Латунь', '1/2"', 1, 0, 0),
    ('Гигиенический душ Iddis', 'Душевые системы', 'Iddis', 6900, 8200,
     'Гигиенический душ с держателем и шлангом. Матовый чёрный.',
     '🖤', 'from-gray-700 to-gray-900', 'Латунь', '1/2"', 1, 0, 1),
    ('Унитаз подвесной Roca Inspira', 'Унитазы', 'Roca', 32900, None,
     'Подвесной унитаз с системой Rimless. Размер 36×56 см.',
     '🚽', 'from-slate-500 to-gray-700', 'Керамика', '—', 1, 0, 1),
    ('Унитаз-компакт Vitra S50', 'Унитазы', 'Vitra', 21500, 24900,
     'Напольный унитаз-компакт с микролифтом и двухрежимным смывом.',
     '🚽', 'from-gray-400 to-slate-600', 'Керамика', '—', 1, 0, 0),
    ('Раковина накладная Roca Uno', 'Раковины', 'Roca', 12400, None,
     'Накладная раковина 45×35 см. Белая керамика.',
     '🪣', 'from-blue-300 to-cyan-400', 'Керамика', '—', 1, 0, 0),
    ('Раковина с тумбой Am.Pm', 'Раковины', 'Am.Pm', 28900, 33900,
     'Комплект: раковина + тумба с ящиками. Подвесная, белый глянец.',
     '🪞', 'from-teal-400 to-emerald-500', 'МДФ/Керамика', '—', 1, 1, 1),
    ('Ванна акриловая Roca Uno 170', 'Ванны', 'Roca', 34900, None,
     'Акриловая ванна 170×75 см. Усиленная, с каркасом.',
     '🛁', 'from-cyan-400 to-blue-500', 'Акрил', '—', 1, 0, 0),
    ('Ванна чугунная 150 см', 'Ванны', 'Прочие', 55900, 62900,
     'Классическая чугунная ванна с эмалью. Долговечная.',
     '🛁', 'from-gray-500 to-slate-700', 'Чугун', '—', 1, 0, 0),
    ('Полотенцесушитель водяной Лесенка', 'Полотенцесушители', 'Прочие', 8900, None,
     'Водяной полотенцесушитель 50×50 см. Нержавейка.',
     '🔥', 'from-orange-400 to-red-500', 'Нержавеющая сталь', '1"', 1, 0, 0),
    ('Полотенцесушитель электрический', 'Полотенцесушители', 'Прочие', 12900, 14900,
     'Электрический с терморегулятором. 500 Вт.',
     '♨️', 'from-red-400 to-pink-500', 'Нержавеющая сталь', '—', 1, 1, 0),
]


def init_db():
    db = get_db()
    db.executescript(SCHEMA)
    if db.execute('SELECT COUNT(*) FROM products').fetchone()[0] == 0:
        db.executemany(
            '''INSERT INTO products
               (name, category, brand, price, old_price, description, icon, color,
                material, diameter, in_stock, is_new, is_hit)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', PRODUCTS)
    db.commit()


with app.app_context():
    init_db()


# ================== ХЕЛПЕРЫ ==================
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            flash('Войдите в аккаунт, чтобы продолжить', 'warning')
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return wrapper


def get_cart():
    return session.get('cart', {})


def save_cart(cart):
    session['cart'] = cart
    session.modified = True


def get_cart_items():
    cart = get_cart()
    if not cart:
        return [], 0
    db = get_db()
    items, total = [], 0
    for pid, qty in cart.items():
        p = db.execute('SELECT * FROM products WHERE id=?', (pid,)).fetchone()
        if p:
            sub = p['price'] * qty
            total += sub
            items.append({'product': p, 'quantity': qty, 'subtotal': sub})
    return items, total


@app.context_processor
def inject_globals():
    user = None
    if 'user_id' in session:
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone()
    return {
        'cart_count': sum(get_cart().values()),
        'current_user': user,
        'categories': ['Смесители', 'Душевые системы', 'Унитазы',
                       'Раковины', 'Ванны', 'Полотенцесушители'],
        'brands_list': ['Grohe', 'Hansgrohe', 'Roca', 'Vitra',
                        'Iddis', 'Am.Pm', 'Прочие'],
    }


# ================== ГЛАВНАЯ ==================
@app.route('/')
def index():
    db = get_db()
    hits = db.execute('SELECT * FROM products WHERE is_hit=1 LIMIT 8').fetchall()
    new = db.execute('SELECT * FROM products WHERE is_new=1 LIMIT 4').fetchall()
    return render_template('index.html', hits=hits, new=new)


# ================== НОВЫЕ РАЗДЕЛЫ ==================
@app.route('/collections')
def collections():
    db = get_db()
    # Собираем товары по категориям — каждая категория = "коллекция"
    cols = {}
    for c in ['Смесители', 'Душевые системы', 'Унитазы', 'Раковины', 'Ванны', 'Полотенцесушители']:
        cols[c] = db.execute('SELECT * FROM products WHERE category=? LIMIT 4', (c,)).fetchall()
    return render_template('collections.html', collections=cols)


@app.route('/partners')
def partners():
    return render_template('partners.html')


@app.route('/where-to-buy')
def where_to_buy():
    return render_template('where_to_buy.html')


# ================== КАТАЛОГ + ФИЛЬТРЫ ==================
@app.route('/catalog')
def catalog():
    db = get_db()
    category = request.args.get('category', '')
    brand = request.args.get('brand', '')
    price_min = request.args.get('price_min', type=float)
    price_max = request.args.get('price_max', type=float)
    material = request.args.get('material', '')
    sort = request.args.get('sort', 'popular')
    q = request.args.get('q', '').strip()

    sql = 'SELECT * FROM products WHERE 1=1'
    params = []
    if category:
        sql += ' AND category=?'; params.append(category)
    if brand:
        sql += ' AND brand=?'; params.append(brand)
    if material:
        sql += ' AND material=?'; params.append(material)
    if price_min is not None:
        sql += ' AND price>=?'; params.append(price_min)
    if price_max is not None:
        sql += ' AND price<=?'; params.append(price_max)
    if q:
        sql += ' AND (name LIKE ? OR description LIKE ?)'
        params += [f'%{q}%', f'%{q}%']

    sql += {
        'price_asc': ' ORDER BY price ASC',
        'price_desc': ' ORDER BY price DESC',
        'new': ' ORDER BY is_new DESC, id DESC',
    }.get(sort, ' ORDER BY is_hit DESC, id DESC')

    products = db.execute(sql, params).fetchall()
    materials = [r['material'] for r in db.execute(
        'SELECT DISTINCT material FROM products WHERE material != ""').fetchall()]
    return render_template('catalog.html', products=products, materials=materials,
                           category=category, brand=brand, material=material,
                           price_min=price_min, price_max=price_max,
                           sort=sort, q=q)


# ================== ТОВАР ==================
@app.route('/product/<int:pid>')
def product(pid):
    db = get_db()
    p = db.execute('SELECT * FROM products WHERE id=?', (pid,)).fetchone()
    if not p:
        flash('Товар не найден', 'error')
        return redirect(url_for('catalog'))
    similar = db.execute(
        'SELECT * FROM products WHERE category=? AND id!=? LIMIT 4',
        (p['category'], pid)).fetchall()
    in_fav = False
    if 'user_id' in session:
        in_fav = db.execute(
            'SELECT 1 FROM favorites WHERE user_id=? AND product_id=?',
            (session['user_id'], pid)).fetchone() is not None
    return render_template('product.html', p=p, similar=similar, in_fav=in_fav)


# ================== КОРЗИНА ==================
@app.route('/cart')
def cart():
    items, total = get_cart_items()
    return render_template('cart.html', items=items, total=total)


@app.route('/cart/add/<int:pid>', methods=['POST'])
def cart_add(pid):
    qty = int(request.form.get('qty', 1))
    cart = get_cart()
    cart[str(pid)] = cart.get(str(pid), 0) + qty
    save_cart(cart)
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return jsonify({'ok': True, 'count': sum(cart.values())})
    flash('Товар добавлен в корзину', 'success')
    return redirect(request.referrer or url_for('cart'))


@app.route('/cart/update', methods=['POST'])
def cart_update():
    pid = request.form.get('pid')
    qty = int(request.form.get('qty', 1))
    cart = get_cart()
    if qty <= 0:
        cart.pop(pid, None)
    else:
        cart[pid] = qty
    save_cart(cart)
    return redirect(url_for('cart'))


@app.route('/cart/remove/<pid>', methods=['POST'])
def cart_remove(pid):
    cart = get_cart()
    cart.pop(str(pid), None)
    save_cart(cart)
    return redirect(url_for('cart'))


# ================== ОФОРМЛЕНИЕ ==================
@app.route('/checkout', methods=['GET', 'POST'])
@login_required
def checkout():
    items, total = get_cart_items()
    if not items:
        flash('Корзина пуста', 'warning')
        return redirect(url_for('catalog'))

    if request.method == 'POST':
        address = request.form.get('address', '').strip()
        phone = request.form.get('phone', '').strip()
        comment = request.form.get('comment', '').strip()
        if not address or not phone:
            flash('Заполните адрес и телефон', 'error')
        else:
            db = get_db()
            cur = db.execute(
                'INSERT INTO orders (user_id, total, address, phone, comment) VALUES (?,?,?,?,?)',
                (session['user_id'], total, address, phone, comment))
            order_id = cur.lastrowid
            for it in items:
                db.execute(
                    'INSERT INTO order_items (order_id, product_id, product_name, quantity, price) VALUES (?,?,?,?,?)',
                    (order_id, it['product']['id'], it['product']['name'],
                     it['quantity'], it['product']['price']))
            db.commit()
            session['cart'] = {}
            flash(f'Заказ #{order_id} оформлен! Мы свяжемся с вами.', 'success')
            return redirect(url_for('account'))

    return render_template('checkout.html', items=items, total=total)


# ================== АВТОРИЗАЦИЯ ==================
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        name = request.form.get('name', '').strip()
        if not email or not password:
            flash('Заполните все поля', 'error')
        elif len(password) < 6:
            flash('Пароль минимум 6 символов', 'error')
        else:
            db = get_db()
            try:
                db.execute('INSERT INTO users (email, password_hash, name) VALUES (?,?,?)',
                           (email, generate_password_hash(password), name))
                db.commit()
                user = db.execute('SELECT * FROM users WHERE email=?', (email,)).fetchone()
                session['user_id'] = user['id']
                flash(f'Добро пожаловать, {name or email}!', 'success')
                return redirect(url_for('index'))
            except sqlite3.IntegrityError:
                flash('Такой email уже зарегистрирован', 'error')
    return render_template('register.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        db = get_db()
        user = db.execute('SELECT * FROM users WHERE email=?', (email,)).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            flash('Вы вошли в аккаунт', 'success')
            nxt = request.args.get('next') or url_for('index')
            return redirect(nxt)
        flash('Неверный email или пароль', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('user_id', None)
    flash('Вы вышли из аккаунта', 'info')
    return redirect(url_for('index'))


# ================== ЛИЧНЫЙ КАБИНЕТ ==================
@app.route('/account')
@login_required
def account():
    db = get_db()
    orders = db.execute(
        'SELECT * FROM orders WHERE user_id=? ORDER BY id DESC',
        (session['user_id'],)).fetchall()
    favs = db.execute(
        '''SELECT p.* FROM favorites f JOIN products p ON p.id=f.product_id
           WHERE f.user_id=?''', (session['user_id'],)).fetchall()
    return render_template('account.html', orders=orders, favorites=favs)


@app.route('/account/profile', methods=['POST'])
@login_required
def profile_update():
    name = request.form.get('name', '').strip()
    phone = request.form.get('phone', '').strip()
    address = request.form.get('address', '').strip()
    db = get_db()
    db.execute('UPDATE users SET name=?, phone=?, address=? WHERE id=?',
               (name, phone, address, session['user_id']))
    db.commit()
    flash('Профиль обновлён', 'success')
    return redirect(url_for('account'))


@app.route('/favorites/toggle/<int:pid>', methods=['POST'])
@login_required
def fav_toggle(pid):
    db = get_db()
    row = db.execute('SELECT 1 FROM favorites WHERE user_id=? AND product_id=?',
                     (session['user_id'], pid)).fetchone()
    if row:
        db.execute('DELETE FROM favorites WHERE user_id=? AND product_id=?',
                   (session['user_id'], pid))
        state = False
    else:
        db.execute('INSERT INTO favorites (user_id, product_id) VALUES (?,?)',
                   (session['user_id'], pid))
        state = True
    db.commit()
    return jsonify({'ok': True, 'in_fav': state})


# ================== КОНТАКТЫ ==================
@app.route('/contacts')
def contacts():
    return render_template('contacts.html')


# ================== API ==================
@app.route('/api/cart/count')
def api_cart_count():
    return jsonify({'count': sum(get_cart().values())})


if __name__ == '__main__':
    app.run(debug=True, port=5000)