from flask import Flask, render_template, redirect, session, request, send_file
import sqlite3
import os
import random
import smtplib
from io import BytesIO
from email.message import EmailMessage

from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from dotenv import load_dotenv

import psycopg2
from psycopg2.extras import DictCursor

import cloudinary
import cloudinary.uploader

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "change-this-secret-key")

UPLOAD_FOLDER = "static/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "1234")

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_APP_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")

DEFAULT_COUPONS = {
    "SAVE10": 10,
}

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True
)


def is_postgres():
    return bool(os.getenv("DATABASE_URL"))


def fix_sql(sql):
    """Convert SQLite ? placeholders to PostgreSQL %s placeholders."""
    if is_postgres():
        return sql.replace("?", "%s")
    return sql


def get_db():
    database_url = os.getenv("DATABASE_URL")

    if database_url:
        return psycopg2.connect(database_url, cursor_factory=DictCursor)

    conn = sqlite3.connect("grocery.db")
    conn.row_factory = sqlite3.Row
    return conn


def execute(cursor, sql, params=()):
    cursor.execute(fix_sql(sql), params)


def fetch_product_id(product):
    return product["id"] if hasattr(product, "keys") and "id" in product.keys() else product[0]


def product_col(product, key, index, default=""):
    try:
        if hasattr(product, "keys") and key in product.keys():
            return product[key]
    except Exception:
        pass

    try:
        return product[index]
    except Exception:
        return default


def product_stock(product):
    try:
        return int(product_col(product, "stock", 5, 0) or 0)
    except (TypeError, ValueError):
        return 0


def can_add_to_cart(product_id, current_quantity):
    conn = get_db()
    cursor = conn.cursor()
    execute(cursor, "SELECT * FROM products WHERE id=?", (product_id,))
    product = cursor.fetchone()
    conn.close()

    if not product:
        return False

    return product_stock(product) > current_quantity


def coupon_value(coupon):
    try:
        if hasattr(coupon, "keys") and "discount" in coupon.keys():
            return int(coupon["discount"] or 0)
    except Exception:
        pass

    try:
        return int(coupon[0] or 0)
    except Exception:
        return 0


def order_col(order, key, index, default=""):
    try:
        if hasattr(order, "keys") and key in order.keys():
            return order[key]
    except Exception:
        pass

    try:
        return order[index]
    except Exception:
        return default


def init_db():
    conn = get_db()
    cursor = conn.cursor()

    if is_postgres():
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id SERIAL PRIMARY KEY,
                name TEXT,
                category TEXT,
                price INTEGER,
                image TEXT,
                stock INTEGER,
                description TEXT,
                image2 TEXT,
                image3 TEXT,
                image4 TEXT,
                image5 TEXT,
                discount INTEGER DEFAULT 10
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS reviews (
                id SERIAL PRIMARY KEY,
                product_id INTEGER,
                customer_name TEXT,
                rating INTEGER,
                comment TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS wishlist (
                id SERIAL PRIMARY KEY,
                product_id INTEGER UNIQUE
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                name TEXT,
                email TEXT UNIQUE,
                password TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id SERIAL PRIMARY KEY,
                customer_name TEXT,
                phone TEXT,
                address TEXT,
                total INTEGER,
                payment_method TEXT,
                payment_status TEXT,
                gst_amount INTEGER,
                status TEXT,
                customer_email TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS order_items (
                id SERIAL PRIMARY KEY,
                order_id INTEGER,
                product_name TEXT,
                price INTEGER,
                quantity INTEGER
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS coupons (
                id SERIAL PRIMARY KEY,
                code TEXT UNIQUE,
                discount INTEGER
            )
        """)

        # Add columns safely for old deployed databases.
        extra_columns = {
            "products": [
                ("description", "TEXT"),
                ("image2", "TEXT"),
                ("image3", "TEXT"),
                ("image4", "TEXT"),
                ("image5", "TEXT"),
                ("discount", "INTEGER DEFAULT 10"),
            ],
            "orders": [
                ("status", "TEXT"),
                ("customer_email", "TEXT"),
            ],
        }

        for table, columns in extra_columns.items():
            for name, col_type in columns:
                try:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN {name} {col_type}")
                except Exception:
                    conn.rollback()
                    cursor = conn.cursor()

    else:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                category TEXT,
                price INTEGER,
                image TEXT,
                stock INTEGER,
                description TEXT,
                image2 TEXT,
                image3 TEXT,
                image4 TEXT,
                image5 TEXT,
                discount INTEGER DEFAULT 10
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS reviews (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER,
                customer_name TEXT,
                rating INTEGER,
                comment TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS wishlist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER UNIQUE
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                email TEXT UNIQUE,
                password TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_name TEXT,
                phone TEXT,
                address TEXT,
                total INTEGER,
                payment_method TEXT,
                payment_status TEXT,
                gst_amount INTEGER,
                status TEXT,
                customer_email TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER,
                product_name TEXT,
                price INTEGER,
                quantity INTEGER
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS coupons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT UNIQUE,
                discount INTEGER
            )
        """)

        extra_columns = {
            "products": [
                ("description", "TEXT"),
                ("image2", "TEXT"),
                ("image3", "TEXT"),
                ("image4", "TEXT"),
                ("image5", "TEXT"),
                ("discount", "INTEGER DEFAULT 10"),
            ],
            "orders": [
                ("status", "TEXT"),
                ("customer_email", "TEXT"),
            ],
        }

        for table, columns in extra_columns.items():
            for name, col_type in columns:
                try:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN {name} {col_type}")
                except sqlite3.OperationalError:
                    pass

    conn.commit()
    conn.close()


def ensure_product_columns():
    init_db()


init_db()


def send_email(to_email, subject, body):
    if not EMAIL_ADDRESS or not EMAIL_APP_PASSWORD:
        print("Email settings missing")
        return

    try:
        msg = EmailMessage()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.set_content(body)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        print("Email send error:", e)


@app.route("/")
def home():
    search = request.args.get("search", "")
    category = request.args.get("category", "")

    conn = get_db()
    cursor = conn.cursor()

    query = "SELECT * FROM products WHERE 1=1"
    params = []

    if search:
        query += " AND name LIKE ?"
        params.append("%" + search + "%")

    if category:
        query += " AND category=?"
        params.append(category)

    execute(cursor, query, params)
    products = cursor.fetchall()

    ratings = {}

    for product in products:
        product_id = fetch_product_id(product)
        execute(
            cursor,
            "SELECT AVG(rating), COUNT(*) FROM reviews WHERE product_id=?",
            (product_id,)
        )
        data = cursor.fetchone()

        avg_rating = round(data[0], 1) if data and data[0] else 0
        total_reviews = data[1] if data else 0

        ratings[product_id] = {
            "avg": avg_rating,
            "count": total_reviews
        }

    cursor.execute("SELECT DISTINCT category FROM products")
    categories = cursor.fetchall()

    cursor.execute("SELECT product_id FROM wishlist")
    wishlist_ids = {row[0] for row in cursor.fetchall()}

    conn.close()

    return render_template(
        "index.html",
        products=products,
        ratings=ratings,
        search=search,
        category=category,
        categories=categories,
        wishlist_ids=wishlist_ids
    )


@app.route("/send_otp", methods=["GET", "POST"])
def send_otp():
    if request.method == "POST":
        email = request.form["email"]
        otp = str(random.randint(100000, 999999))

        session["otp_email"] = email
        session["otp"] = otp

        send_email(
            email,
            "Your Damodar Grocery Store OTP",
            f"Your OTP is {otp}"
        )

        return redirect("/verify_otp")

    return render_template("send_otp.html")


@app.route("/verify_otp", methods=["GET", "POST"])
def verify_otp():
    if request.method == "POST":
        user_otp = request.form["otp"]

        if user_otp == session.get("otp"):
            session["user"] = session.get("otp_email")
            session.pop("otp", None)
            return redirect("/")

        return "Wrong OTP"

    return render_template("verify_otp.html")


@app.route("/apply_coupon", methods=["POST"])
def apply_coupon():
    code = request.form["coupon"].upper().strip()

    if not code:
        session["coupon_code"] = ""
        session["discount"] = 0
        session["coupon_message"] = "Please enter a coupon code."
        session["coupon_status"] = "error"
        return redirect("/cart")

    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT discount FROM coupons WHERE code=?", (code,))
    coupon = cursor.fetchone()

    conn.close()

    if coupon:
        session["coupon_code"] = code
        session["discount"] = coupon_value(coupon)
        session["coupon_message"] = f"Coupon {code} applied successfully."
        session["coupon_status"] = "success"
    elif code in DEFAULT_COUPONS:
        session["coupon_code"] = code
        session["discount"] = DEFAULT_COUPONS[code]
        session["coupon_message"] = f"Coupon {code} applied successfully."
        session["coupon_status"] = "success"
    else:
        session["coupon_code"] = ""
        session["discount"] = 0
        session["coupon_message"] = "Invalid coupon code."
        session["coupon_status"] = "error"

    return redirect("/cart")


@app.route("/review/<int:product_id>", methods=["GET", "POST"])
def review(product_id):
    if request.method == "POST":
        customer_name = request.form["customer_name"]
        rating = request.form["rating"]
        comment = request.form["comment"]

        conn = get_db()
        cursor = conn.cursor()

        execute(
            cursor,
            "INSERT INTO reviews(product_id, customer_name, rating, comment) VALUES(?,?,?,?)",
            (product_id, customer_name, rating, comment)
        )

        conn.commit()
        conn.close()

        return redirect("/reviews/" + str(product_id))

    return render_template("review.html", product_id=product_id)


@app.route("/reviews/<int:product_id>")
def reviews(product_id):
    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT * FROM reviews WHERE product_id=?", (product_id,))
    reviews_data = cursor.fetchall()

    conn.close()

    return render_template("reviews.html", reviews=reviews_data)


@app.route("/profile")
def profile():
    if "user" not in session:
        return redirect("/customer_login")
    return render_template("profile.html", name=session["user"])


@app.route("/my_orders")
def my_orders():
    if "user" not in session:
        return redirect("/customer_login")

    conn = get_db()
    cursor = conn.cursor()

    user_email_or_name = session["user"]
    execute(
        cursor,
        "SELECT * FROM orders WHERE customer_email=? OR customer_name=? ORDER BY id DESC",
        (user_email_or_name, user_email_or_name)
    )
    orders = cursor.fetchall()

    conn.close()

    return render_template("my_orders.html", orders=orders)


@app.route("/admin_login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if request.form["username"] == ADMIN_USERNAME and request.form["password"] == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect("/orders")
        return "Wrong admin username or password"

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form["name"]
        email = request.form["email"]
        password = request.form["password"]

        hashed_password = generate_password_hash(password)

        conn = get_db()
        cursor = conn.cursor()

        try:
            execute(
                cursor,
                "INSERT INTO users(name, email, password) VALUES(?,?,?)",
                (name, email, hashed_password)
            )
            conn.commit()
        except Exception:
            conn.rollback()
            conn.close()
            return "Email already registered"

        conn.close()
        return redirect("/customer_login")

    return render_template("register.html")


@app.route("/customer_login", methods=["GET", "POST"])
def customer_login():
    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]

        conn = get_db()
        cursor = conn.cursor()
        execute(cursor, "SELECT * FROM users WHERE email=?", (email,))
        user = cursor.fetchone()
        conn.close()

        if user:
            stored_password = user["password"] if hasattr(user, "keys") and "password" in user.keys() else user[3]

            # Supports old plain-text passwords and upgrades them after login.
            if check_password_hash(stored_password, password) or stored_password == password:
                session["user"] = email

                if stored_password == password:
                    conn = get_db()
                    cursor = conn.cursor()
                    execute(
                        cursor,
                        "UPDATE users SET password=? WHERE email=?",
                        (generate_password_hash(password), email)
                    )
                    conn.commit()
                    conn.close()

                return redirect("/")

        return "Wrong email or password"

    return render_template("customer_login.html")


@app.route("/add_to_cart/<int:product_id>")
def add_to_cart(product_id):
    cart = session.get("cart", {})
    if isinstance(cart, list):
        cart = {}

    cart_key = str(product_id)
    current_quantity = int(cart.get(cart_key, 0))

    if not can_add_to_cart(product_id, current_quantity):
        return redirect("/cart")

    cart[cart_key] = current_quantity + 1

    session["cart"] = cart
    session.modified = True

    return redirect("/cart")


@app.route("/increase/<int:product_id>")
def increase(product_id):
    cart = session.get("cart", {})
    if isinstance(cart, list):
        cart = {}

    cart_key = str(product_id)
    current_quantity = int(cart.get(cart_key, 0))

    if not can_add_to_cart(product_id, current_quantity):
        return redirect("/cart")

    cart[cart_key] = current_quantity + 1

    session["cart"] = cart
    session.modified = True

    return redirect("/cart")


@app.route("/decrease/<int:product_id>")
def decrease(product_id):
    cart = session.get("cart", {})
    if isinstance(cart, list):
        cart = {}

    product_id = str(product_id)

    if product_id in cart:
        cart[product_id] -= 1
        if cart[product_id] <= 0:
            cart.pop(product_id)

    session["cart"] = cart
    session.modified = True

    return redirect("/cart")


@app.route("/cart")
def cart():
    cart_items = session.get("cart", {})
    if isinstance(cart_items, list):
        cart_items = {}
        session["cart"] = {}

    conn = get_db()
    cursor = conn.cursor()

    products = []
    total = 0

    for product_id, quantity in cart_items.items():
        execute(cursor, "SELECT * FROM products WHERE id=?", (product_id,))
        product = cursor.fetchone()

        if product:
            price = int(product_col(product, "price", 3, 0))
            subtotal = price * quantity
            total += subtotal
            products.append((product, quantity, subtotal))

    gst = int(total * 0.05)
    discount_percent = int(session.get("discount", 0) or 0)
    discount_amount = int(total * discount_percent / 100)
    grand_total = total + gst - discount_amount

    conn.close()

    return render_template(
        "cart.html",
        products=products,
        total=total,
        gst=gst,
        discount_percent=discount_percent,
        discount_amount=discount_amount,
        grand_total=grand_total,
        coupon_code=session.get("coupon_code", ""),
        coupon_message=session.get("coupon_message", ""),
        coupon_status=session.get("coupon_status", "")
    )


@app.route("/product/<int:product_id>")
def product_details(product_id):
    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT * FROM products WHERE id=?", (product_id,))
    product = cursor.fetchone()

    if not product:
        conn.close()
        return "Product not found"

    execute(
        cursor,
        "SELECT AVG(rating), COUNT(*) FROM reviews WHERE product_id=?",
        (product_id,)
    )
    data = cursor.fetchone()

    avg_rating = round(data[0], 1) if data and data[0] else 0
    total_reviews = data[1] if data else 0

    execute(cursor, "SELECT * FROM reviews WHERE product_id=? ORDER BY id DESC", (product_id,))
    reviews_data = cursor.fetchall()

    cursor.execute("SELECT product_id FROM wishlist")
    wishlist_ids = {row[0] for row in cursor.fetchall()}

    execute(
        cursor,
        "SELECT * FROM products WHERE category=? AND id!=? LIMIT 5",
        (product_col(product, "category", 2), product_id)
    )
    related_products = cursor.fetchall()

    conn.close()

    return render_template(
        "product_details.html",
        product=product,
        avg_rating=avg_rating,
        total_reviews=total_reviews,
        reviews=reviews_data,
        wishlist_ids=wishlist_ids,
        related_products=related_products
    )


@app.route("/wishlist/<int:product_id>")
def wishlist(product_id):
    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT * FROM wishlist WHERE product_id=?", (product_id,))
    existing = cursor.fetchone()

    if existing:
        execute(cursor, "DELETE FROM wishlist WHERE product_id=?", (product_id,))
    else:
        try:
            execute(cursor, "INSERT INTO wishlist(product_id) VALUES(?)", (product_id,))
        except Exception:
            conn.rollback()

    conn.commit()
    conn.close()

    return redirect("/")


@app.route("/wishlist")
def wishlist_page():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT products.*
        FROM wishlist
        JOIN products ON wishlist.product_id = products.id
    """)

    products = cursor.fetchall()
    conn.close()

    return render_template("wishlist.html", products=products)


@app.route("/checkout", methods=["GET", "POST"])
def checkout():
    if request.method == "POST":
        name = request.form["name"]
        phone = request.form["phone"]
        address = request.form["address"]
        payment_method = request.form["payment_method"]
        email = request.form["email"]

        cart_items = session.get("cart", {})
        if isinstance(cart_items, list):
            cart_items = {}

        if not cart_items:
            return redirect("/cart")

        conn = get_db()
        cursor = conn.cursor()

        total = 0
        order_products = []

        for product_id, quantity in cart_items.items():
            execute(cursor, "SELECT * FROM products WHERE id=?", (product_id,))
            product = cursor.fetchone()

            if product:
                stock = product_stock(product)
                price = int(product_col(product, "price", 3, 0))

                if stock < quantity:
                    conn.close()
                    return f"Not enough stock for {product_col(product, 'name', 1)}"

                subtotal = price * quantity
                total += subtotal
                order_products.append((product, quantity))

        gst = int(total * 0.05)
        discount_percent = int(session.get("discount", 0) or 0)
        discount_amount = int(total * discount_percent / 100)
        grand_total = total + gst - discount_amount

        if is_postgres():
            execute(
                cursor,
                """
                INSERT INTO orders(customer_name, phone, address, total, payment_method, payment_status, gst_amount, status, customer_email)
                VALUES(?,?,?,?,?,?,?,?,?)
                RETURNING id
                """,
                (
                    name,
                    phone,
                    address,
                    grand_total,
                    payment_method,
                    "Paid" if payment_method == "Demo Payment" else "Pending",
                    gst,
                    "Pending",
                    email
                )
            )
            order_id = cursor.fetchone()[0]
        else:
            execute(
                cursor,
                """
                INSERT INTO orders(customer_name, phone, address, total, payment_method, payment_status, gst_amount, status, customer_email)
                VALUES(?,?,?,?,?,?,?,?,?)
                """,
                (
                    name,
                    phone,
                    address,
                    grand_total,
                    payment_method,
                    "Paid" if payment_method == "Demo Payment" else "Pending",
                    gst,
                    "Pending",
                    email
                )
            )
            order_id = cursor.lastrowid

        for product, quantity in order_products:
            product_id = fetch_product_id(product)

            execute(
                cursor,
                "UPDATE products SET stock = stock - ? WHERE id=? AND stock >= ?",
                (quantity, product_id, quantity)
            )

            execute(
                cursor,
                "INSERT INTO order_items(order_id, product_name, price, quantity) VALUES(?,?,?,?)",
                (
                    order_id,
                    product_col(product, "name", 1),
                    product_col(product, "price", 3),
                    quantity
                )
            )

        conn.commit()
        conn.close()

        send_email(
            email,
            "Order Placed Successfully",
            f"Your order #{order_id} has been placed. Total amount: Rs.{grand_total}"
        )

        session["cart"] = {}
        session["discount"] = 0
        session["coupon_code"] = ""
        session["last_order_id"] = order_id

        return render_template("success.html", order_id=order_id)

    return render_template("checkout.html")


@app.route("/invoice/<int:order_id>")
def invoice(order_id):
    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT * FROM orders WHERE id=?", (order_id,))
    order = cursor.fetchone()

    execute(cursor, "SELECT * FROM order_items WHERE order_id=?", (order_id,))
    items = cursor.fetchall()

    conn.close()

    if not order:
        return "Order not found"

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer)
    pdf.drawString(100, 800, "Damodar Grocery Store Invoice")
    pdf.drawString(100, 770, f"Order ID: {order_col(order, 'id', 0)}")
    pdf.drawString(100, 750, f"Customer: {order_col(order, 'customer_name', 1)}")
    pdf.drawString(100, 730, f"Phone: {order_col(order, 'phone', 2)}")
    pdf.drawString(100, 710, f"Address: {order_col(order, 'address', 3)}")

    y = 670

    for item in items:
        item_name = order_col(item, "product_name", 2)
        item_price = order_col(item, "price", 3)
        item_qty = order_col(item, "quantity", 4)
        pdf.drawString(100, y, f"{item_name} | Price: Rs.{item_price} | Qty: {item_qty}")
        y -= 20

    pdf.drawString(100, y - 20, f"Grand Total: Rs.{order_col(order, 'total', 4)}")
    pdf.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"invoice_{order_id}.pdf",
        mimetype="application/pdf"
    )


@app.route("/orders")
def orders():
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM orders ORDER BY id DESC")
    all_orders = cursor.fetchall()

    total_orders = len(all_orders)
    total_sales = sum(int(order_col(order, "total", 4, 0) or 0) for order in all_orders)

    pending_orders = 0
    delivered_orders = 0

    for order in all_orders:
        if order_col(order, "status", 8) == "Delivered":
            delivered_orders += 1
        else:
            pending_orders += 1

    conn.close()

    return render_template(
        "orders.html",
        orders=all_orders,
        total_orders=total_orders,
        total_sales=total_sales,
        pending_orders=pending_orders,
        delivered_orders=delivered_orders
    )


@app.route("/update_status/<int:order_id>/<status>")
def update_status(order_id, status):
    if "admin" not in session:
        return redirect("/admin_login")

    allowed_status = ["Pending", "Packed", "Shipped", "Delivered", "Cancelled"]
    if status not in allowed_status:
        return "Invalid status"

    conn = get_db()
    cursor = conn.cursor()
    execute(cursor, "UPDATE orders SET status=? WHERE id=?", (status, order_id))
    conn.commit()
    conn.close()

    return redirect("/orders")


@app.route("/export_orders")
def export_orders():
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM orders ORDER BY id DESC")
    orders = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append(["ID", "Name", "Phone", "Address", "Total", "Payment Method", "Payment Status", "GST", "Status", "Email"])

    for order in orders:
        ws.append([
            order_col(order, "id", 0),
            order_col(order, "customer_name", 1),
            order_col(order, "phone", 2),
            order_col(order, "address", 3),
            order_col(order, "total", 4),
            order_col(order, "payment_method", 5),
            order_col(order, "payment_status", 6),
            order_col(order, "gst_amount", 7),
            order_col(order, "status", 8),
            order_col(order, "customer_email", 9),
        ])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="orders_export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/add_product", methods=["GET", "POST"])
def add_product():
    if "admin" not in session:
        return redirect("/admin_login")

    ensure_product_columns()

    if request.method == "POST":
        name = request.form["name"]
        category = request.form["category"]
        price = request.form["price"]
        stock = request.form["stock"]
        description = request.form.get("description", "")

        image_files = request.files.getlist("images")
        images = []

        for file in image_files[:5]:
            if file and file.filename:
                secure_filename(file.filename)
                upload_result = cloudinary.uploader.upload(file)
                images.append(upload_result["secure_url"])

        image1 = images[0] if len(images) > 0 else ""
        image2 = images[1] if len(images) > 1 else ""
        image3 = images[2] if len(images) > 2 else ""
        image4 = images[3] if len(images) > 3 else ""
        image5 = images[4] if len(images) > 4 else ""

        conn = get_db()
        cursor = conn.cursor()

        execute(
            cursor,
            """
            INSERT INTO products
            (name, category, price, image, stock, description, image2, image3, image4, image5)
            VALUES (?,?,?,?,?,?,?,?,?,?)
            """,
            (
                name,
                category,
                price,
                image1,
                stock,
                description,
                image2,
                image3,
                image4,
                image5
            )
        )

        conn.commit()
        conn.close()

        return redirect("/")

    return render_template("add_product.html")


@app.route("/edit_product/<int:product_id>", methods=["GET", "POST"])
def edit_product(product_id):
    if "admin" not in session:
        return redirect("/admin_login")

    ensure_product_columns()

    conn = get_db()
    cursor = conn.cursor()

    execute(cursor, "SELECT * FROM products WHERE id=?", (product_id,))
    product = cursor.fetchone()

    if not product:
        conn.close()
        return "Product not found"

    if request.method == "POST":
        name = request.form.get("name", "")
        category = request.form.get("category", "")
        price = request.form.get("price", "")
        stock = request.form.get("stock", "")
        description = request.form.get("description", "")

        old_images = [
            product_col(product, "image", 4, ""),
            product_col(product, "image2", 7, ""),
            product_col(product, "image3", 8, ""),
            product_col(product, "image4", 9, ""),
            product_col(product, "image5", 10, ""),
        ]

        fields = ["image1", "image2", "image3", "image4", "image5"]
        new_images = []

        for index, field_name in enumerate(fields):
            file = request.files.get(field_name)

            if file and file.filename:
                secure_filename(file.filename)
                upload_result = cloudinary.uploader.upload(file)
                new_images.append(upload_result["secure_url"])
            else:
                new_images.append(old_images[index])

        execute(
            cursor,
            """
            UPDATE products
            SET name=?,
                category=?,
                price=?,
                image=?,
                stock=?,
                description=?,
                image2=?,
                image3=?,
                image4=?,
                image5=?
            WHERE id=?
            """,
            (
                name,
                category,
                price,
                new_images[0],
                stock,
                description,
                new_images[1],
                new_images[2],
                new_images[3],
                new_images[4],
                product_id
            )
        )

        conn.commit()

        execute(cursor, "SELECT description FROM products WHERE id=?", (product_id,))
        saved_description = cursor.fetchone()

        conn.close()

        if saved_description and saved_description[0] == description:
            return redirect("/product/" + str(product_id))

        return "Description save nahi hua. Database column problem hai."

    conn.close()
    return render_template("edit_product.html", product=product)


@app.route("/delete_product/<int:product_id>")
def delete_product(product_id):
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()
    execute(cursor, "DELETE FROM products WHERE id=?", (product_id,))
    conn.commit()
    conn.close()

    return redirect("/")


if __name__ == "__main__":
    app.run(debug=True)
