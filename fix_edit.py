from pathlib import Path

path = Path("app.py")
text = path.read_text(encoding="utf-8")

start = text.index('@app.route("/edit_product/<int:product_id>", methods=["GET", "POST"])')
end = text.index('@app.route("/delete_product/<int:product_id>")')

new_code = '''@app.route("/edit_product/<int:product_id>", methods=["GET", "POST"])
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
        name = request.form["name"]
        category = request.form["category"]
        price = request.form["price"]
        stock = request.form["stock"]
        description = request.form.get("description", "")

        image1 = product[4]
        image2 = product[8] if len(product) > 8 else ""
        image3 = product[9] if len(product) > 9 else ""
        image4 = product[10] if len(product) > 10 else ""
        image5 = product[11] if len(product) > 11 else ""

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
                file.save(os.path.join(UPLOAD_FOLDER, filename))
                new_images.append("/" + UPLOAD_FOLDER + "/" + filename)
            else:
                new_images.append(old_image)

        cursor.execute(
            """
            UPDATE products
            SET name=?, category=?, price=?, image=?, stock=?, description=?,
                image2=?, image3=?, image4=?, image5=?
            WHERE id=?
            """,
            (
                name, category, price, new_images[0], stock, description,
                new_images[1], new_images[2], new_images[3], new_images[4],
                product_id
            )
        )

        conn.commit()
        conn.close()

        return redirect("/")

    conn.close()
    return render_template("edit_product.html", product=product)


'''

text = text[:start] + new_code + text[end:]
path.write_text(text, encoding="utf-8")

print("app.py fixed successfully")