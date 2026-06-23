from flask import Flask, render_template, redirect, session, request, send_file
import sqlite3
import os
import random
import smtplib
from email.message import EmailMessage
from werkzeug.utils import secure_filename
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from dotenv import load_dotenv

import psycopg2
from psycopg2.extras import DictCursor

import cloudinary
import cloudinary.uploader

load_dotenv()

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True
)

app = Flask(__name__)
app.secret_key = "damodar123"


UPLOAD_FOLDER = "static/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "1234"

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_APP_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")


def get_db():
    database_url = os.getenv("DATABASE_URL")

    if database_url:
        return psycopg2.connect(database_url, cursor_factory=DictCursor)

    conn = sqlite3.connect("grocery.db")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cursor = conn.cursor()

    if os.getenv("DATABASE_URL"):
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
                product_id INTEGER
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
                status TEXT
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

    else:
        columns = [
            ("description", "TEXT"),
            ("image2", "TEXT"),
            ("image3", "TEXT"),
            ("image4", "TEXT"),
            ("image5", "TEXT"),
            ("discount", "INTEGER DEFAULT 10")
        ]

        for name, column_type in columns:
            try:
                cursor.execute(f"ALTER TABLE products ADD COLUMN {name} {column_type}")
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

    msg = EmailMessage()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        smtp.send_message(msg)


@app.route("/")
def home():

    search = request.args.get("search", "")
    category = request.args.get("category", "")

    conn = get_db()
    cursor = conn.cursor()

    is_postgres = bool(os.getenv("DATABASE_URL"))

    query = "SELECT * FROM products WHERE 1=1"
    params = []

    if search:
        query += " AND name LIKE %s" if is_postgres else " AND name LIKE ?"
        params.append("%" + search + "%")

    if category:
        query += " AND category=%s" if is_postgres else " AND category=?"
        params.append(category)

    cursor.execute(query, params)
    products = cursor.fetchall()

    ratings = {}

    for product in products:
        cursor.execute(
            "SELECT AVG(rating), COUNT(*) FROM reviews WHERE product_id=%s" if is_postgres else
            "SELECT AVG(rating), COUNT(*) FROM reviews WHERE product_id=?",
            (product[0],)
        )
        data = cursor.fetchone()

        avg_rating = round(data[0], 1) if data[0] else 0
        total_reviews = data[1]

        ratings[product[0]] = {
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
        else:
            return "Wrong OTP"

    return render_template("verify_otp.html")


@app.route("/apply_coupon", methods=["POST"])
def apply_coupon():
    code = request.form["coupon"].upper()

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT discount FROM coupons WHERE code=?", (code,))
    coupon = cursor.fetchone()

    conn.close()

    if coupon:
        session["coupon_code"] = code
        session["discount"] = coupon[0]
    else:
        session["coupon_code"] = ""
        session["discount"] = 0

    return redirect("/cart")


@app.route("/review/<int:product_id>", methods=["GET", "POST"])
def review(product_id):
    if request.method == "POST":
        customer_name = request.form["customer_name"]
        rating = request.form["rating"]
        comment = request.form["comment"]

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute(
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

    cursor.execute("SELECT * FROM reviews WHERE product_id=?", (product_id,))
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
    cursor.execute("SELECT * FROM orders ORDER BY id DESC")
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

        conn = get_db()
        cursor = conn.cursor()

        try:
            cursor.execute(
                "INSERT INTO users(name, email, password) VALUES(?,?,?)",
                (name, email, password)
            )
            conn.commit()
        except sqlite3.IntegrityError:
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
        cursor.execute(
            "SELECT * FROM users WHERE email=? AND password=?",
            (email, password)
        )
        user = cursor.fetchone()
        conn.close()

        if user:
            session["user"] = user[1]
            return redirect("/")

        return "Wrong email or password"

    return render_template("customer_login.html")


@app.route("/add_to_cart/<int:product_id>")
def add_to_cart(product_id):
    cart = session.get("cart", {})
    if isinstance(cart, list):
        cart = {}

    product_id = str(product_id)
    cart[product_id] = cart.get(product_id, 0) + 1

    session["cart"] = cart
    session.modified = True

    return redirect("/cart")


@app.route("/increase/<int:product_id>")
def increase(product_id):
    cart = session.get("cart", {})
    if isinstance(cart, list):
        cart = {}

    product_id = str(product_id)
    cart[product_id] = cart.get(product_id, 0) + 1

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
        cursor.execute("SELECT * FROM products WHERE id=?", (product_id,))
        product = cursor.fetchone()

        if product:
            subtotal = int(product[3]) * quantity
            total += subtotal
            products.append((product, quantity, subtotal))

    gst = int(total * 0.05)
    discount_percent = session.get("discount", 0)
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
        grand_total=grand_total
    )

@app.route("/product/<int:product_id>")
def product_details(product_id):
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM products WHERE id=?", (product_id,))
    product = cursor.fetchone()

    if not product:
        conn.close()
        return "Product not found"

    cursor.execute(
        "SELECT AVG(rating), COUNT(*) FROM reviews WHERE product_id=?",
        (product_id,)
    )
    data = cursor.fetchone()

    avg_rating = round(data[0], 1) if data[0] else 0
    total_reviews = data[1]

    cursor.execute("SELECT * FROM reviews WHERE product_id=? ORDER BY id DESC", (product_id,))
    reviews_data = cursor.fetchall()

    cursor.execute("SELECT product_id FROM wishlist")
    wishlist_ids = {row[0] for row in cursor.fetchall()}

    cursor.execute(
        "SELECT * FROM products WHERE category=? AND id!=? LIMIT 5",
        (product[2], product_id)
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

    cursor.execute("SELECT * FROM wishlist WHERE product_id=?", (product_id,))
    existing = cursor.fetchone()

    if existing:
        cursor.execute("DELETE FROM wishlist WHERE product_id=?", (product_id,))
    else:
        cursor.execute("INSERT INTO wishlist(product_id) VALUES(?)", (product_id,))

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

        conn = get_db()
        cursor = conn.cursor()

        total = 0
        order_products = []

        for product_id, quantity in cart_items.items():
            cursor.execute("SELECT * FROM products WHERE id=?", (product_id,))
            product = cursor.fetchone()

            if product:
                subtotal = int(product[3]) * quantity
                total += subtotal
                order_products.append((product, quantity))

                cursor.execute(
                    "UPDATE products SET stock = stock - ? WHERE id=? AND stock >= ?",
                    (quantity, product_id, quantity)
                )

        gst = int(total * 0.05)
        discount_percent = session.get("discount", 0)
        discount_amount = int(total * discount_percent / 100)
        grand_total = total + gst - discount_amount

        cursor.execute(
            """
            INSERT INTO orders(customer_name, phone, address, total, payment_method, payment_status, gst_amount, status)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            (
                name,
                phone,
                address,
                grand_total,
                payment_method,
                "Paid" if payment_method == "Demo Payment" else "Pending",
                gst,
                "Pending"
            )
        )

        order_id = cursor.lastrowid

        for product, quantity in order_products:
            cursor.execute(
                "INSERT INTO order_items(order_id, product_name, price, quantity) VALUES(?,?,?,?)",
                (order_id, product[1], product[3], quantity)
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
    filename = f"invoice_{order_id}.pdf"

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM orders WHERE id=?", (order_id,))
    order = cursor.fetchone()

    cursor.execute("SELECT * FROM order_items WHERE order_id=?", (order_id,))
    items = cursor.fetchall()

    conn.close()

    if not order:
        return "Order not found"

    pdf = canvas.Canvas(filename)
    pdf.drawString(100, 800, "Damodar Grocery Store Invoice")
    pdf.drawString(100, 770, f"Order ID: {order[0]}")
    pdf.drawString(100, 750, f"Customer: {order[1]}")
    pdf.drawString(100, 730, f"Phone: {order[2]}")
    pdf.drawString(100, 710, f"Address: {order[3]}")

    y = 670

    for item in items:
        pdf.drawString(100, y, f"{item[2]} | Price: Rs.{item[3]} | Qty: {item[4]}")
        y -= 20

    pdf.drawString(100, y - 20, f"Grand Total: Rs.{order[4]}")
    pdf.save()

    return send_file(filename, as_attachment=True)


@app.route("/orders")
def orders():
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM orders")
    all_orders = cursor.fetchall()

    total_orders = len(all_orders)
    total_sales = sum(order[4] for order in all_orders)

    pending_orders = 0
    delivered_orders = 0

    for order in all_orders:
        if len(order) > 8 and order[8] == "Delivered":
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

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE orders SET status=? WHERE id=?", (status, order_id))
    conn.commit()
    conn.close()

    return redirect("/orders")


@app.route("/export_orders")
def export_orders():
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM orders")
    orders = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    ws.append(["ID", "Name", "Phone", "Address", "Total"])

    for order in orders:
        ws.append([order[0], order[1], order[2], order[3], order[4]])

    filename = "orders_export.xlsx"
    wb.save(filename)

    return send_file(filename, as_attachment=True)


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

        for file in image_files:
            if file and file.filename:
                filename = secure_filename(file.filename)
                upload_result = cloudinary.uploader.upload(file)
                images.append(upload_result["secure_url"])

        image1 = images[0] if len(images) > 0 else ""
        image2 = images[1] if len(images) > 1 else ""
        image3 = images[2] if len(images) > 2 else ""
        image4 = images[3] if len(images) > 3 else ""
        image5 = images[4] if len(images) > 4 else ""

        conn = get_db()
        cursor = conn.cursor()

        cursor.execute(
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

    cursor.execute("SELECT * FROM products WHERE id=?", (product_id,))
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

        image1 = product[4] if len(product) > 4 else ""
        image2 = product[7] if len(product) > 7 else ""
        image3 = product[8] if len(product) > 8 else ""
        image4 = product[9] if len(product) > 9 else ""
        image5 = product[10] if len(product) > 10 else ""

        files = [
            ("image1", image1),
            ("image2", image2),
            ("image3", image3),
            ("image4", image4),
            ("image5", image5)
        ]

        new_images = []

        for field_name, old_image in files:
            file = request.files.get(field_name)

            if file and file.filename:
                filename = secure_filename(file.filename)
                upload_result = cloudinary.uploader.upload(file)
                new_images.append(upload_result["secure_url"])
            else:
                new_images.append(old_image)

        cursor.execute(
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

        cursor.execute("SELECT description FROM products WHERE id=?", (product_id,))
        saved_description = cursor.fetchone()

        conn.close()

        if saved_description and saved_description[0] == description:
            return redirect("/product/" + str(product_id))
        else:
            return "Description save nahi hua. Database column problem hai."

    conn.close()
    return render_template("edit_product.html", product=product)

@app.route("/delete_product/<int:product_id>")
def delete_product(product_id):
    if "admin" not in session:
        return redirect("/admin_login")

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM products WHERE id=?", (product_id,))
    conn.commit()
    conn.close()

    return redirect("/")


if __name__ == "__main__":
    app.run(debug=True)