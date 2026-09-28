from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = "far-dimsum-2026"

DB = "far_dimsum.db"

# =========================
# LOGIN
# =========================

USERS = {
    "asyifa": "1123",
    "muhsafar": "2311",
    "admin": "1234"
}

# =========================
# SEMUA MENU FAR DIMSUM
# =========================

MENU = {
    "Dimsum Mentai": {
        "Mentai Isi 3": 15000,
        "Mentai Isi 4": 20000,
        "Mentai Isi 6": 25000,
        "Mentai Isi 8": 35000,
        "Mentai Isi 12": 50000
    },

    "Mentai Keju": {
        "Mentai Keju Isi 3": 18000,
        "Mentai Keju Isi 4": 23000,
        "Mentai Keju Isi 6": 28000,
        "Mentai Keju Isi 8": 38000,
        "Mentai Keju Isi 12": 56000
    },

    "Dimsum Mix": {
        "Mix Isi 4": 23000,
        "Mix Isi 6": 28000,
        "Mix Isi 8": 38000,
        "Mix Isi 12": 56000
    },

    "Dimsum Bakar": {
        "Bakar Original Isi 3": 15000,
        "Bakar Original Isi 4": 18000,
        "Bakar Lava Isi 3": 15000,
        "Bakar Lava Isi 4": 18000,
        "Bakar Mentai Isi 3": 15000,
        "Bakar Mentai Isi 4": 20000,
        "Bakar Mentai Keju Isi 3": 18000,
        "Bakar Mentai Keju Isi 4": 23000
    },

    "Special Paket": {
        "Dimsum Birthday": 65000
    }
}


# =========================
# DATABASE
# =========================

def koneksi():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def buat_database():
    conn = koneksi()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT NOT NULL,
            kasir TEXT NOT NULL,
            tipe TEXT NOT NULL,
            total INTEGER NOT NULL,
            pembayaran INTEGER NOT NULL,
            kembalian INTEGER NOT NULL,
            detail TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


# =========================
# LOGIN
# =========================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        if username in USERS and USERS[username] == password:

            session["user"] = username

            return redirect(url_for("kasir"))

        return render_template(
            "login.html",
            error="Username atau password salah."
        )

    return render_template("login.html")


@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("login"))


# =========================
# HALAMAN KASIR
# =========================

@app.route("/")
def kasir():

    if "user" not in session:
        return redirect(url_for("login"))

    return render_template(
        "kasir.html",
        menu=MENU,
        kasir=session["user"]
    )


# =========================
# SIMPAN TRANSAKSI
# =========================

@app.route("/simpan_transaksi", methods=["POST"])
def simpan_transaksi():

    if "user" not in session:
        return jsonify({
            "success": False,
            "message": "Belum login."
        })

    data = request.get_json()

    keranjang = data["keranjang"]
    tipe = data["tipe"]
    total = int(data["total"])
    pembayaran = int(data["pembayaran"])

    if pembayaran < total:

        return jsonify({
            "success": False,
            "message": "Uang pembayaran kurang."
        })

    kembalian = pembayaran - total

    detail = []

    for item in keranjang:

        detail.append(
            f'{item["nama"]} x{item["jumlah"]}'
        )

    detail = ", ".join(detail)

    tanggal = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn = koneksi()

    cursor = conn.execute("""
        INSERT INTO transaksi
        (
            tanggal,
            kasir,
            tipe,
            total,
            pembayaran,
            kembalian,
            detail
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        tanggal,
        session["user"],
        tipe,
        total,
        pembayaran,
        kembalian,
        detail
    ))

    conn.commit()

    nomor = cursor.lastrowid

    conn.close()

    return jsonify({
        "success": True,
        "nomor": nomor,
        "kembalian": kembalian
    })


# =========================
# RIWAYAT
# =========================

@app.route("/riwayat")
def riwayat():

    if "user" not in session:
        return redirect(url_for("login"))

    tanggal = request.args.get("tanggal")

    if not tanggal:

        tanggal = datetime.now().strftime("%Y-%m-%d")

    conn = koneksi()

    transaksi = conn.execute("""
        SELECT *
        FROM transaksi
        WHERE tanggal LIKE ?
        ORDER BY id DESC
    """, (
        tanggal + "%",
    )).fetchall()

    conn.close()

    total_omzet = sum(
        row["total"]
        for row in transaksi
    )

    return render_template(
        "riwayat.html",
        transaksi=transaksi,
        tanggal=tanggal,
        total_omzet=total_omzet
    )


# =========================
# EXPORT EXCEL
# =========================

@app.route("/excel")
def excel():

    if "user" not in session:
        return redirect(url_for("login"))

    tanggal = request.args.get("tanggal")

    if not tanggal:
        tanggal = datetime.now().strftime("%Y-%m-%d")

    conn = koneksi()

    transaksi = conn.execute("""
        SELECT *
        FROM transaksi
        WHERE tanggal LIKE ?
        ORDER BY id
    """, (
        tanggal + "%",
    )).fetchall()

    conn.close()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Riwayat Transaksi"

    sheet.append([
        "ID",
        "Tanggal",
        "Kasir",
        "Tipe",
        "Total",
        "Pembayaran",
        "Kembalian",
        "Detail"
    ])

    for row in transaksi:

        sheet.append([
            row["id"],
            row["tanggal"],
            row["kasir"],
            row["tipe"],
            row["total"],
            row["pembayaran"],
            row["kembalian"],
            row["detail"]
        ])

    file = BytesIO()

    workbook.save(file)

    file.seek(0)

    return app.response_class(
        file.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
            f"attachment; filename=Far_Dimsum_{tanggal}.xlsx"
        }
    )


# =========================
# START
# =========================

if __name__ == "__main__":

    buat_database()

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
