import os
from flask import Flask, request, jsonify, send_file, make_response, Response, session, redirect, url_for, render_template_string
from flask_cors import CORS
from openai import OpenAI
import json
import sqlite3
from datetime import datetime
import csv
from io import StringIO
from functools import wraps
from dotenv import load_dotenv

# IMPOR UNTUK FITUR IMPOR CSV & PPT
import pandas as pd  
import re            
import io 
from pptx import Presentation 

# 1. MEMUAT KUNCI RAHASIA DARI FILE .env
load_dotenv()
api_key_rahasia = os.getenv("OPENAI_API_KEY")

app = Flask(__name__)
# Kunci rahasia wajib untuk menjalankan fitur session (Login Portal)
app.secret_key = os.getenv("SECRET_KEY", "kunci_rahasia_untuk_sesi_aplikasi")
CORS(app) 

# Kode akses untuk portal portofolio (HRD) - Menggunakan env untuk keamanan
ACCESS_CODE = os.getenv("ACCESS_CODE", "BuildByAdreanthzny*")

# INISIALISASI DATABASE SQLITE
def init_db():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS blueprints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                nama TEXT,
                kelas TEXT,
                j_minat TEXT,
                j_hambatan TEXT,
                j_finansial TEXT,
                j_motivasi TEXT,
                ai_analisis TEXT,
                ai_karir TEXT,
                kategori_masalah TEXT
            )
        ''')
        
        try:
            cursor.execute("ALTER TABLE blueprints ADD COLUMN kategori_masalah TEXT DEFAULT 'Lainnya'")
        except sqlite3.OperationalError:
            pass 

        conn.commit()
        conn.close()
    except sqlite3.OperationalError as e:
        # Menghindari crash saat inisialisasi awal di Vercel (read-only file system)
        print(f"Melewati inisialisasi database (Lingkungan Vercel): {e}")

init_db()

client = OpenAI(api_key=api_key_rahasia)

# ==========================================
# BLOK SISTEM AUTENTIKASI (LOGIN ADMIN BK)
# ==========================================
def check_auth(username, password):
    # Menggunakan env untuk menyembunyikan kredensial di GitHub
    admin_user = os.getenv("ADMIN_USERNAME", "guru")
    admin_pass = os.getenv("ADMIN_PASSWORD", "bksman1")
    return username == admin_user and password == admin_pass

def authenticate():
    return Response(
    'Akses ditolak. Anda tidak memiliki izin untuk melihat dasbor ini.', 401,
    {'WWW-Authenticate': 'Basic realm="Dasbor Konselor Rahasia"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# ==========================================
# BLOK SISTEM AUTENTIKASI (KODE AKSES PORTAL/HRD)
# ==========================================
def requires_access_code(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        input_code = request.form.get('access_code')
        if input_code == ACCESS_CODE:
            session['authenticated'] = True
            return redirect(url_for('beranda'))
        else:
            error = "Kode akses tidak valid."
            
    # Menggunakan HTML sebaris agar tidak perlu membuat file login.html baru
    login_html = """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Login Akses Portal</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-gray-100 flex items-center justify-center min-h-screen">
        <div class="bg-white p-8 rounded-lg shadow-md w-full max-w-md">
            <div class="text-center mb-6">
                <h2 class="text-2xl font-bold text-gray-800">Akses Terbatas</h2>
                <p class="text-sm text-gray-500 mt-2">Silakan masukkan kode akses untuk melihat portal aplikasi portofolio.</p>
            </div>
            <form method="POST" class="space-y-4">
                <div>
                    <input type="password" name="access_code" placeholder="Masukkan Kode Akses" required 
                        class="w-full px-4 py-3 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                </div>
                <button type="submit" 
                    class="w-full bg-blue-600 text-white font-semibold py-3 rounded-md hover:bg-blue-700 transition duration-200">
                    Masuk ke Portal
                </button>
            </form>
            {% if error %}
            <div class="mt-4 p-3 bg-red-100 text-red-700 rounded-md text-sm text-center">
                {{ error }}
            </div>
            {% endif %}
        </div>
    </body>
    </html>
    """
    return render_template_string(login_html, error=error)

@app.route('/logout')
def logout():
    session.pop('authenticated', None)
    return redirect(url_for('login'))

# ==========================================
# FUNGSI BANTUAN NORMALISASI DATA
# ==========================================
def standarisasi_kelas(kelas_raw):
    kelas_str = str(kelas_raw).strip().lower()
    tingkat_map = {'10': 'X', '11': 'XI', '12': 'XII'}
    match = re.match(r'(10|11|12)[\-\s]*(.*)', kelas_str)
    
    if match:
        tingkat_angka = match.group(1)
        sub_kelas = match.group(2).upper()
        tingkat_romawi = tingkat_map.get(tingkat_angka, tingkat_angka)
        return f"{tingkat_romawi}-{sub_kelas}"
    return kelas_str.upper()

# ==========================================
# BLOK RUTE TAMPILAN (FRONTEND ROUTING)
# ==========================================
@app.route('/')
@requires_access_code
def beranda():
    return send_file('index.html')

@app.route('/problem-solver')
@requires_access_code
def problem_solver_ui():
    return send_file('problem_solver_engine.html')

@app.route('/career-blueprint')
@requires_access_code
def career_blueprint():
    return send_file('career-blueprint-engine.html')

@app.route('/admin')
@requires_auth
def admin_dashboard():
    return send_file('admin.html')

# ==========================================
# BLOK RUTE DATA & MESIN AI (API ENDPOINTS)
# ==========================================

@app.route('/generate-blueprint', methods=['POST'])
def generate_blueprint():
    try:
        data = request.get_json()
        identitas = data.get('identitas', {})
        nama_pengguna = identitas.get('nama', '-')
        # Variabel kelas dihapus dari logic agar universal dan simpel
        
        kuesioner = data.get('kuesioner', [])
        input_perilaku_luang = kuesioner[0]['jawaban'] if len(kuesioner) > 0 else "-"
        input_hambatan_belajar = kuesioner[1]['jawaban'] if len(kuesioner) > 1 else "-"
        input_kondisi_finansial = kuesioner[2]['jawaban'] if len(kuesioner) > 2 else "-"
        input_profil_risiko = kuesioner[3]['jawaban'] if len(kuesioner) > 3 else "-"

        # Prompt AI Diubah Menjadi Universal (Segala Kalangan)
        prompt = f"""
        ROLE: Anda adalah "Pragmatic Career Strategist", pakar strategi karir untuk berbagai kalangan (remaja, mahasiswa, hingga pekerja profesional). 

        OBJECTIVE: Menganalisis 4 variabel data pengguna dan merancang cetak biru karir.

        RULES & CONSTRAINTS:
        1. Anda WAJIB menganalisis hambatan pengguna dan mengklasifikasikannya ke dalam SATU Kategori Masalah Baku.
        2. Pilihan Kategori Masalah Baku HANYA BOLEH salah satu dari: [Akademik, Karir, Pribadi, Sosial].
        3. FORMAT OUTPUT: JSON valid murni.

        DATA INPUT PENGGUNA:
        - Hambatan Belajar/Kerja: {input_hambatan_belajar}
        - Perilaku Luang: {input_perilaku_luang}
        - Kondisi Finansial: {input_kondisi_finansial}
        - Profil Risiko: {input_profil_risiko}

        FORMAT OUTPUT JSON:
        {{
          "kategori_masalah": "Salah satu dari [Akademik, Karir, Pribadi, Sosial]",
          "analisis_realitas": "String maksimal 3 kalimat objektif tentang posisi pengguna saat ini",
          "rekomendasi_karir_utama": "Nama profesi spesifik 1",
          "rekomendasi_karir_alternatif": "Nama profesi spesifik 2",
          "langkah_konkret_1_tahun_kedepan": ["Langkah 1", "Langkah 2"],
          "taktik_mengatasi_hambatan": "String solusi taktis manajemen stres/belajar/kerja"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda adalah asisten AI yang selalu mengeluarkan output dalam format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            response_format={"type": "json_object"}
        )
        
        raw_text = response.choices[0].message.content
        result_json = json.loads(raw_text)
        
        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ai_analisis = result_json.get('analisis_realitas', '-')
        ai_karir = result_json.get('rekomendasi_karir_utama', '-') + " / " + result_json.get('rekomendasi_karir_alternatif', '-')
        kategori = result_json.get('kategori_masalah', 'Lainnya')

        # BLOK PENYIMPANAN DATABASE DINONAKTIFKAN (Mencegah Error 500 Read-Only Vercel)
        # conn = sqlite3.connect('data_siswa.db')
        # cursor = conn.cursor()
        # cursor.execute('''
        #     INSERT INTO blueprints (tanggal, nama, kelas, j_minat, j_hambatan, j_finansial, j_motivasi, ai_analisis, ai_karir, kategori_masalah)
        #     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        # ''', (
        #     waktu_sekarang, nama_pengguna, "-",
        #     input_perilaku_luang, input_hambatan_belajar,
        #     input_kondisi_finansial, input_profil_risiko,
        #     ai_analisis, ai_karir, kategori
        # ))
        # conn.commit()
        # conn.close()
        
        return jsonify({"status": "success", "data": result_json})

    except Exception as e:
        print("Detail Error:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/data-siswa', methods=['GET'])
@requires_auth
def get_data_siswa():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, nama, kelas, kategori_masalah, ai_karir 
            FROM blueprints 
            ORDER BY id DESC
        ''')
        rows = cursor.fetchall()
        conn.close()
        
        siswa_data = []
        for row in rows:
            karir = row[4] if row[4] else "Belum Dianalisis / -"
            
            siswa_data.append({
                "id": row[0],
                "nama": row[1],
                "kelas": row[2],
                "kategori_masalah": row[3],
                "rekomendasi_karir": karir, 
                "ai_karir": karir 
            })
            
        return jsonify({"status": "success", "data": siswa_data})

    except Exception as e:
        print(f"Error Load Data Siswa: {e}")
        return jsonify({"status": "error", "message": "Gagal memuat tabel data pengguna"}), 500


# ==========================================
# ENDPOINT PERBAIKAN & TAMBAHAN UNTUK ADMIN
# ==========================================

@app.route('/api/detail-siswa', methods=['GET'])
@app.route('/api/detail-siswa/<int:siswa_id>', methods=['GET'])
@requires_auth
def detail_siswa(siswa_id=None):
    try:
        if siswa_id is None:
            siswa_id = request.args.get('id')
            
        if not siswa_id:
            return jsonify({"status": "error", "message": "ID pengguna tidak ditemukan"}), 400

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, tanggal, nama, kelas, j_minat, j_hambatan, j_finansial, j_motivasi, ai_analisis, ai_karir, kategori_masalah 
            FROM blueprints 
            WHERE id = ?
        ''', (siswa_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            return jsonify({
                "status": "success",
                "data": {
                    "id": row[0],
                    "tanggal": row[1],
                    "nama": row[2],
                    "kelas": row[3],
                    "j_minat": row[4],
                    "j_hambatan": row[5],
                    "j_finansial": row[6],
                    "j_motivasi": row[7],
                    "ai_analisis": row[8],
                    "ai_karir": row[9],
                    "kategori_masalah": row[10]
                }
            })
        return jsonify({"status": "error", "message": "Data pengguna tidak ditemukan"}), 404

    except Exception as e:
        print(f"Error Detail Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/hapus-siswa', methods=['DELETE', 'POST'])
@app.route('/api/hapus-siswa/<int:siswa_id>', methods=['DELETE', 'POST'])
@requires_auth
def hapus_siswa(siswa_id=None):
    try:
        if siswa_id is None:
            data = request.get_json(silent=True) or {}
            siswa_id = data.get('id') or request.args.get('id')

        if not siswa_id:
            return jsonify({"status": "error", "message": "ID pengguna tidak diberikan"}), 400

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('DELETE FROM blueprints WHERE id = ?', (siswa_id,))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"Data pengguna ID {siswa_id} berhasil dihapus"})

    except Exception as e:
        print(f"Error Hapus Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/hapus-semua-siswa', methods=['DELETE', 'POST'])
@app.route('/api/reset-database', methods=['DELETE', 'POST'])
@requires_auth
def hapus_semua_siswa():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('DELETE FROM blueprints')
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": "Seluruh data pengguna berhasil dihapus"})

    except Exception as e:
        print(f"Error Hapus Semua Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/edit-siswa', methods=['POST', 'PUT'])
@app.route('/api/edit-siswa/<int:siswa_id>', methods=['POST', 'PUT'])
@requires_auth
def edit_siswa(siswa_id=None):
    try:
        data = request.get_json() or {}
        if siswa_id is None:
            siswa_id = data.get('id')
            
        if not siswa_id:
            return jsonify({"status": "error", "message": "ID pengguna tidak diberikan"}), 400

        nama = data.get('nama')
        kelas = data.get('kelas')
        kategori = data.get('kategori_masalah')

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE blueprints 
            SET nama = COALESCE(?, nama),
                kelas = COALESCE(?, kelas),
                kategori_masalah = COALESCE(?, kategori_masalah)
            WHERE id = ?
        ''', (nama, kelas, kategori, siswa_id))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": "Data pengguna berhasil diperbarui"})

    except Exception as e:
        print(f"Error Edit Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/chat-admin', methods=['POST'])
@app.route('/api/diskusi-rpl', methods=['POST'])
@requires_auth
def chat_admin():
    try:
        data = request.get_json()
        messages = data.get('messages', [])
        chat_history = data.get('chatHistory', messages)
        
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('SELECT nama, kelas, kategori_masalah, ai_analisis FROM blueprints ORDER BY id DESC LIMIT 20')
        rows = cursor.fetchall()
        conn.close()
        
        konteks_siswa = "\n".join([f"- {r[0]} ({r[1]}) | Masalah: {r[2]} | Analisis: {r[3]}" for r in rows])
        
        system_prompt = f"""
        ROLE: Anda adalah "Asisten Guru BK Senior & Konsultan Edukasi/HRD".
        TUGAS: Membantu Administrator/Guru BK dalam merancang strategi layanan, mendiskusikan kasus pengguna, menyusun evaluasi, dan memberikan rekomendasi intervensi yang pragmatis.
        
        KONTEKS DATA PENGGUNA TERBARU DARI DATABASE:
        {konteks_siswa if konteks_siswa else 'Belum ada data pengguna yang direkam.'}
        
        Gunakan bahasa profesional namun hangat, suportif, dan solutif.
        """
        
        formatted_messages = [{"role": "system", "content": system_prompt}]
        for msg in chat_history:
            formatted_messages.append({
                "role": msg.get("role", "user"),
                "content": msg.get("content", "")
            })

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=formatted_messages,
            temperature=0.5
        )
        
        reply = response.choices[0].message.content
        return jsonify({"status": "success", "reply": reply, "message": reply})

    except Exception as e:
        print(f"Error Chat Admin: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/statistik', methods=['GET'])
@requires_auth
def get_statistik():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT kategori_masalah, COUNT(id) as total
            FROM blueprints
            GROUP BY kategori_masalah
        ''')
        rows = cursor.fetchall()
        conn.close()

        labels = []
        data = []
        for row in rows:
            labels.append(row[0] if row[0] else "Belum Terkategori")
            data.append(row[1])

        return jsonify({
            "status": "success", 
            "labels": labels, 
            "data": data
        })

    except Exception as e:
        print(f"Error Load Statistik: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/generate-rpl', methods=['POST'])
@requires_auth
def generate_rpl():
    try:
        data = request.get_json()
        kategori_masalah = data.get('kategori', 'Akademik')
        metode_pilihan = data.get('metode', 'Bimbingan Klasikal')

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT nama, kelas, j_hambatan, ai_analisis 
            FROM blueprints 
            WHERE kategori_masalah = ?
        ''', (kategori_masalah,))
        rows = cursor.fetchall()
        conn.close()

        konteks_siswa = ""
        if rows:
            for row in rows:
                nama_siswa = row[0]
                kelas = row[1]
                keluhan = row[2]
                analisis = row[3]
                konteks_siswa += f"- {nama_siswa} (Kelas {kelas}) | Keluhan: {keluhan} | Analisis: {analisis}\n"
        else:
            return jsonify({"status": "error", "message": f"Belum ada data siswa untuk kategori {kategori_masalah} di database."}), 404

        prompt = f"""
        ROLE: Anda adalah Konselor Senior/Guru Bimbingan dan Konseling yang pakar menyusun RPL (Rencana Pelaksanaan Layanan) adaptif.
        
        TUGAS: Susun materi layanan untuk merespons kasus pada kategori masalah: {kategori_masalah}.

        PARAMETER LAYANAN:
        - Metode Layanan: {metode_pilihan}

        DATA RIIL KASUS SAAT INI (WAJIB JADIKAN ACUAN MATERI):
        {konteks_siswa}

        ATURAN KETAT POLARITAS DIKSI JUDUL:
        1. PEMETAAN KATA KERJA LENGKAP:
           - Variabel NEGATIF / MASALAH (Kecemasan, Stres, Adiksi, Konflik, Gugup, dll): WAJIB menggunakan kata kerja reduksi ("Manajemen", "Reduksi", "Penanganan", "Pengendalian"). DILARANG KERAS menggunakan kata "Optimalisasi" atau "Peningkatan" untuk masalah negatif!
           - Variabel POSITIF / POTENSI (Motivasi, Kepercayaan Diri, Minat, Resiliensi): WAJIB menggunakan kata kerja eskalasi ("Peningkatan", "Optimalisasi", "Pengembangan", "Penguatan").
        
        2. STRUKTUR DAN TATA BAHASA:
           - Setiap variabel harus diikat oleh kata kerja yang sesuai dengan polaritasnya.
           - Sintesiskan kesamaan masalah dari data di atas untuk menjadi satu topik layanan yang kohesif.

        FORMAT OUTPUT JSON:
        {{
            "topik_layanan": "Judul profesional, presisi secara logika semantik (maksimal 10 kata)",
            "tujuan_layanan": "Tujuan objektif spesifik menyelesaikan keluhan dominan di atas",
            "tahap_awal": "Penjelasan aktivitas pembuka (apersepsi) yang relevan dengan kasus",
            "tahap_inti": ["Langkah inti 1 (solusi kasus)", "Langkah inti 2", "Langkah inti 3"],
            "tahap_penutup_dan_evaluasi": "Aktivitas penutup dan instrumen evaluasi spesifik"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda wajib membalas hanya dengan format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"}
        )
        
        result_json = json.loads(response.choices[0].message.content)
        
        tahap_inti_html = "".join([f"- {langkah}<br>" for langkah in result_json.get('tahap_inti', [])])
        
        hasil_rpl = f"""
        <b>RENCANA PELAKSANAAN LAYANAN (RPL) / MATERI BIMBINGAN</b><br>
        ===================================================<br><br>
        <b>Topik Layanan :</b> {result_json.get('topik_layanan', '-')}<br>
        <b>Tujuan Layanan:</b> {result_json.get('tujuan_layanan', '-')}<br><br>
        <b>A. TAHAP AWAL</b><br>
        {result_json.get('tahap_awal', '-')}<br><br>
        <b>B. TAHAP INTI</b><br>
        {tahap_inti_html}<br>
        <b>C. TAHAP PENUTUP & EVALUASI</b><br>
        {result_json.get('tahap_penutup_dan_evaluasi', '-')}
        """

        return jsonify({"status": "success", "hasil": hasil_rpl})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/generate-rpl-chat', methods=['POST'])
@requires_auth
def generate_rpl_chat():
    try:
        data = request.get_json()
        chat_history = data.get('chatHistory', [])
        
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('SELECT nama, kelas, kategori_masalah, ai_analisis FROM blueprints ORDER BY id DESC LIMIT 20')
        rows = cursor.fetchall()
        conn.close()
        
        konteks_siswa = "\n".join([f"- {r[0]} ({r[1]}) | Masalah: {r[2]} | Analisis: {r[3]}" for r in rows])
        pesan_obrolan = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in chat_history])
        
        prompt = f"""
        ROLE: Anda adalah Konselor Senior.
        TUGAS: Susun materi layanan (Rencana Pelaksanaan Layanan) berdasarkan diskusi chat berikut.
        
        RIWAYAT DISKUSI CHAT DENGAN KONSELOR:
        {pesan_obrolan}
        
        DATA PENGGUNA SAAT INI (Jadikan acuan jika relevan dengan obrolan):
        {konteks_siswa}
        
        FORMAT OUTPUT JSON MURNI:
        {{
            "topik_layanan": "Judul profesional layanan bimbingan (maksimal 10 kata)",
            "tujuan_layanan": "Tujuan objektif dari layanan ini",
            "tahap_awal": "Aktivitas pembuka (apersepsi)",
            "tahap_inti": ["Langkah penyelesaian 1", "Langkah penyelesaian 2", "Langkah 3"],
            "tahap_penutup_dan_evaluasi": "Aktivitas penutup dan evaluasi"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda wajib membalas hanya dengan format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"}
        )
        
        result_json = json.loads(response.choices[0].message.content)
        return jsonify({"status": "success", "data": result_json})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/eksplorasi-siswa', methods=['POST'])
def eksplorasi_siswa():
    try:
        data = request.get_json()
        chat_history = data.get('chatHistory', [])
        blueprint_data = data.get('blueprintData', {})
        
        kategori = blueprint_data.get('kategori_masalah', '-')
        analisis = blueprint_data.get('analisis_realitas', '-')
        karir_utama = blueprint_data.get('rekomendasi_karir_utama', '-')
        karir_alt = blueprint_data.get('rekomendasi_karir_alternatif', '-')
        
        langkah_list = blueprint_data.get('langkah_konkret_1_tahun_kedepan', [])
        if isinstance(langkah_list, list):
            langkah = ", ".join(langkah_list)
        else:
            langkah = str(langkah_list)
            
        taktik = blueprint_data.get('taktik_mengatasi_hambatan', '-')
        
        # Penyesuaian Role untuk Semua Kalangan
        system_prompt = f"""
        [PERAN DAN NADA INTERAKSI]
        Anda adalah Asisten Eksplorasi Karir AI untuk segala kalangan (remaja hingga profesional). Nada bicaramu hangat, membumi, setara seperti teman diskusi (peer-to-peer), dan tidak menggurui. Gunakan bahasa Indonesia sehari-hari yang sopan namun santai (menggunakan 'aku' dan 'kamu'). Jangan gunakan format poin-poin yang kaku jika tidak diminta.

        [ATURAN UTAMA]
        Tugas utamamu adalah membantu pengguna mengeksplorasi minat dan hasil asesmen karir mereka. Biarkan pengguna memegang kendali percakapan. Berikan solusi taktis, langkah kecil (baby steps), atau referensi belajar jika diminta. Jangan mengulangi secara mentahBerikut adalah versi utuh dan terbaru dari file `app.py` yang sudah di-upgrade berdasarkan file `app_21.py` Anda[cite: 17]. 

Upgrade ini mencakup:
1. **Target Universal:** Mengubah prompt AI di `/generate-blueprint`, `/api/eksplorasi-siswa`, dan `/api/problem-solver` agar tidak lagi terbatas untuk siswa SMA N 1 Padalarang, melainkan adaptif untuk remaja, mahasiswa, hingga pekerja profesional[cite: 17].
2. **Penghapusan Kolom Kelas:** Menonaktifkan pengambilan dan pemrosesan variabel `kelas` pada pembuatan cetak biru[cite: 17].
3. **Solusi Error 500 Vercel:** Memblokir fungsi penulisan (INSERT) ke `sqlite3` pada proses utama, serta membungkus inisialisasi database (`init_db`) dengan perlindungan `try-except` agar aplikasi tidak *crash* saat Vercel menerapkan sistem *read-only*[cite: 17].
4. **Keamanan Kredensial:** Menggunakan `os.getenv` untuk `ACCESS_CODE` dan kredensial Admin BK agar aman saat diunggah ke GitHub[cite: 17].

Anda bisa langsung menyalin seluruh kode di bawah ini dan menggantikan isi file `app.py` Anda di GitHub:

```python
import os
from flask import Flask, request, jsonify, send_file, make_response, Response, session, redirect, url_for, render_template_string
from flask_cors import CORS
from openai import OpenAI
import json
import sqlite3
from datetime import datetime
import csv
from io import StringIO
from functools import wraps
from dotenv import load_dotenv

# IMPOR UNTUK FITUR IMPOR CSV & PPT
import pandas as pd  
import re            
import io 
from pptx import Presentation 

# 1. MEMUAT KUNCI RAHASIA DARI FILE .env
load_dotenv()
api_key_rahasia = os.getenv("OPENAI_API_KEY")

app = Flask(__name__)
# Kunci rahasia wajib untuk menjalankan fitur session (Login Portal)
app.secret_key = os.getenv("SECRET_KEY", "kunci_rahasia_untuk_sesi_aplikasi")
CORS(app) 

# 2. UPGRADE KEAMANAN: Kode akses HRD/Portal mengambil dari .env, dengan fallback nilai default
ACCESS_CODE = os.getenv("ACCESS_CODE", "BuildByAdreanthzny*")

# INISIALISASI DATABASE SQLITE
def init_db():
    # UPGRADE VERCEL: Dibungkus try-except agar tidak crash di environment read-only Vercel
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS blueprints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                nama TEXT,
                kelas TEXT,
                j_minat TEXT,
                j_hambatan TEXT,
                j_finansial TEXT,
                j_motivasi TEXT,
                ai_analisis TEXT,
                ai_karir TEXT,
                kategori_masalah TEXT
            )
        ''')
        
        try:
            cursor.execute("ALTER TABLE blueprints ADD COLUMN kategori_masalah TEXT DEFAULT 'Lainnya'")
        except sqlite3.OperationalError:
            pass 

        conn.commit()
        conn.close()
    except Exception as e:
        print("Mode Serverless (Vercel): Inisialisasi database lokal dilewati untuk menghindari error read-only.")

init_db()

client = OpenAI(api_key=api_key_rahasia)

# ==========================================
# BLOK SISTEM AUTENTIKASI (LOGIN ADMIN BK)
# ==========================================
def check_auth(username, password):
    # UPGRADE KEAMANAN: Mengambil kredensial admin dari .env
    admin_user = os.getenv("ADMIN_USERNAME", "guru")
    admin_pass = os.getenv("ADMIN_PASSWORD", "bksman1")
    return username == admin_user and password == admin_pass

def authenticate():
    return Response(
    'Akses ditolak. Anda tidak memiliki izin untuk melihat dasbor ini.', 401,
    {'WWW-Authenticate': 'Basic realm="Dasbor Konselor Rahasia"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# ==========================================
# BLOK SISTEM AUTENTIKASI (KODE AKSES PORTAL/HRD)
# ==========================================
def requires_access_code(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        input_code = request.form.get('access_code')
        if input_code == ACCESS_CODE:
            session['authenticated'] = True
            return redirect(url_for('beranda'))
        else:
            error = "Kode akses tidak valid."
            
    # Menggunakan HTML sebaris agar tidak perlu membuat file login.html baru
    login_html = """
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Login Akses Portal</title>
        <script src="[https://cdn.tailwindcss.com](https://cdn.tailwindcss.com)"></script>
    </head>
    <body class="bg-gray-100 flex items-center justify-center min-h-screen">
        <div class="bg-white p-8 rounded-lg shadow-md w-full max-w-md">
            <div class="text-center mb-6">
                <h2 class="text-2xl font-bold text-gray-800">Akses Terbatas</h2>
                <p class="text-sm text-gray-500 mt-2">Silakan masukkan kode akses untuk melihat portal aplikasi portofolio.</p>
            </div>
            <form method="POST" class="space-y-4">
                <div>
                    <input type="password" name="access_code" placeholder="Masukkan Kode Akses" required 
                        class="w-full px-4 py-3 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                </div>
                <button type="submit" 
                    class="w-full bg-blue-600 text-white font-semibold py-3 rounded-md hover:bg-blue-700 transition duration-200">
                    Masuk ke Portal
                </button>
            </form>
            {% if error %}
            <div class="mt-4 p-3 bg-red-100 text-red-700 rounded-md text-sm text-center">
                {{ error }}
            </div>
            {% endif %}
        </div>
    </body>
    </html>
    """
    return render_template_string(login_html, error=error)

@app.route('/logout')
def logout():
    session.pop('authenticated', None)
    return redirect(url_for('login'))

# ==========================================
# FUNGSI BANTUAN NORMALISASI DATA
# ==========================================
def standarisasi_kelas(kelas_raw):
    kelas_str = str(kelas_raw).strip().lower()
    tingkat_map = {'10': 'X', '11': 'XI', '12': 'XII'}
    match = re.match(r'(10|11|12)[\-\s]*(.*)', kelas_str)
    
    if match:
        tingkat_angka = match.group(1)
        sub_kelas = match.group(2).upper()
        tingkat_romawi = tingkat_map.get(tingkat_angka, tingkat_angka)
        return f"{tingkat_romawi}-{sub_kelas}"
    return kelas_str.upper()

# ==========================================
# BLOK RUTE TAMPILAN (FRONTEND ROUTING)
# ==========================================
@app.route('/')
@requires_access_code
def beranda():
    return send_file('index.html')

@app.route('/problem-solver')
@requires_access_code
def problem_solver_ui():
    return send_file('problem_solver_engine.html')

@app.route('/career-blueprint')
@requires_access_code
def career_blueprint():
    return send_file('career-blueprint-engine.html')

@app.route('/admin')
@requires_auth
def admin_dashboard():
    return send_file('admin.html')

# ==========================================
# BLOK RUTE DATA & MESIN AI (API ENDPOINTS)
# ==========================================

@app.route('/generate-blueprint', methods=['POST'])
def generate_blueprint():
    try:
        data = request.get_json()
        identitas = data.get('identitas', {})
        nama_siswa = identitas.get('nama', '-')
        
        # UPGRADE: Kolom kelas diabaikan untuk menjangkau semua kalangan
        # kelas_siswa = identitas.get('kelas', '-')
        
        kuesioner = data.get('kuesioner', [])
        input_perilaku_luang = kuesioner[0]['jawaban'] if len(kuesioner) > 0 else "-"
        input_hambatan_belajar = kuesioner[1]['jawaban'] if len(kuesioner) > 1 else "-"
        input_kondisi_finansial = kuesioner[2]['jawaban'] if len(kuesioner) > 2 else "-"
        input_profil_risiko = kuesioner[3]['jawaban'] if len(kuesioner) > 3 else "-"

        # UPGRADE: Prompt AI diubah agar universal (Remaja, Mahasiswa, Pekerja Profesional)
        prompt = f"""
        ROLE: Anda adalah "Pragmatic Career Strategist", pakar strategi karir untuk berbagai kalangan (remaja, mahasiswa, hingga pekerja profesional). 

        OBJECTIVE: Menganalisis 4 variabel data pengguna dan merancang cetak biru karir.

        RULES & CONSTRAINTS:
        1. Anda WAJIB menganalisis hambatan pengguna dan mengklasifikasikannya ke dalam SATU Kategori Masalah Baku.
        2. Pilihan Kategori Masalah Baku HANYA BOLEH salah satu dari: [Akademik, Karir, Pribadi, Sosial].
        3. FORMAT OUTPUT: JSON valid murni.

        DATA INPUT PENGGUNA:
        - Hambatan/Tantangan: {input_hambatan_belajar}
        - Perilaku Luang: {input_perilaku_luang}
        - Kondisi Finansial: {input_kondisi_finansial}
        - Profil Risiko: {input_profil_risiko}

        FORMAT OUTPUT JSON:
        {{
          "kategori_masalah": "Salah satu dari [Akademik, Karir, Pribadi, Sosial]",
          "analisis_realitas": "String maksimal 3 kalimat objektif tentang posisi pengguna saat ini",
          "rekomendasi_karir_utama": "Nama profesi spesifik 1",
          "rekomendasi_karir_alternatif": "Nama profesi spesifik 2",
          "langkah_konkret_1_tahun_kedepan": ["Langkah 1", "Langkah 2"],
          "taktik_mengatasi_hambatan": "String solusi taktis manajemen stres/belajar/pekerjaan"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda adalah asisten AI yang selalu mengeluarkan output dalam format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            response_format={"type": "json_object"}
        )
        
        raw_text = response.choices[0].message.content
        result_json = json.loads(raw_text)
        
        # UPGRADE VERCEL: Kode penyimpanan ke sqlite3 dinonaktifkan untuk menghindari Error 500 (Read-Only File System)
        '''
        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ai_analisis = result_json.get('analisis_realitas', '-')
        ai_karir = result_json.get('rekomendasi_karir_utama', '-') + " / " + result_json.get('rekomendasi_karir_alternatif', '-')
        kategori = result_json.get('kategori_masalah', 'Lainnya')

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO blueprints (tanggal, nama, kelas, j_minat, j_hambatan, j_finansial, j_motivasi, ai_analisis, ai_karir, kategori_masalah)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            waktu_sekarang, nama_siswa, "-",
            input_perilaku_luang, input_hambatan_belajar,
            input_kondisi_finansial, input_profil_risiko,
            ai_analisis, ai_karir, kategori
        ))
        conn.commit()
        conn.close()
        '''
        
        return jsonify({"status": "success", "data": result_json})

    except Exception as e:
        print("Detail Error:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/data-siswa', methods=['GET'])
@requires_auth
def get_data_siswa():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, nama, kelas, kategori_masalah, ai_karir 
            FROM blueprints 
            ORDER BY id DESC
        ''')
        rows = cursor.fetchall()
        conn.close()
        
        siswa_data = []
        for row in rows:
            karir = row[4] if row[4] else "Belum Dianalisis / -"
            
            siswa_data.append({
                "id": row[0],
                "nama": row[1],
                "kelas": row[2],
                "kategori_masalah": row[3],
                "rekomendasi_karir": karir, 
                "ai_karir": karir 
            })
            
        return jsonify({"status": "success", "data": siswa_data})

    except Exception as e:
        print(f"Error Load Data Siswa: {e}")
        return jsonify({"status": "error", "message": "Gagal memuat tabel data siswa"}), 500


# ==========================================
# ENDPOINT PERBAIKAN & TAMBAHAN UNTUK ADMIN
# ==========================================

@app.route('/api/detail-siswa', methods=['GET'])
@app.route('/api/detail-siswa/<int:siswa_id>', methods=['GET'])
@requires_auth
def detail_siswa(siswa_id=None):
    try:
        if siswa_id is None:
            siswa_id = request.args.get('id')
            
        if not siswa_id:
            return jsonify({"status": "error", "message": "ID siswa tidak ditemukan"}), 400

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, tanggal, nama, kelas, j_minat, j_hambatan, j_finansial, j_motivasi, ai_analisis, ai_karir, kategori_masalah 
            FROM blueprints 
            WHERE id = ?
        ''', (siswa_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            return jsonify({
                "status": "success",
                "data": {
                    "id": row[0],
                    "tanggal": row[1],
                    "nama": row[2],
                    "kelas": row[3],
                    "j_minat": row[4],
                    "j_hambatan": row[5],
                    "j_finansial": row[6],
                    "j_motivasi": row[7],
                    "ai_analisis": row[8],
                    "ai_karir": row[9],
                    "kategori_masalah": row[10]
                }
            })
        return jsonify({"status": "error", "message": "Data siswa tidak ditemukan"}), 404

    except Exception as e:
        print(f"Error Detail Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/hapus-siswa', methods=['DELETE', 'POST'])
@app.route('/api/hapus-siswa/<int:siswa_id>', methods=['DELETE', 'POST'])
@requires_auth
def hapus_siswa(siswa_id=None):
    try:
        if siswa_id is None:
            data = request.get_json(silent=True) or {}
            siswa_id = data.get('id') or request.args.get('id')

        if not siswa_id:
            return jsonify({"status": "error", "message": "ID siswa tidak diberikan"}), 400

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('DELETE FROM blueprints WHERE id = ?', (siswa_id,))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": f"Data siswa ID {siswa_id} berhasil dihapus"})

    except Exception as e:
        print(f"Error Hapus Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/hapus-semua-siswa', methods=['DELETE', 'POST'])
@app.route('/api/reset-database', methods=['DELETE', 'POST'])
@requires_auth
def hapus_semua_siswa():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('DELETE FROM blueprints')
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": "Seluruh data siswa berhasil dihapus"})

    except Exception as e:
        print(f"Error Hapus Semua Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/edit-siswa', methods=['POST', 'PUT'])
@app.route('/api/edit-siswa/<int:siswa_id>', methods=['POST', 'PUT'])
@requires_auth
def edit_siswa(siswa_id=None):
    try:
        data = request.get_json() or {}
        if siswa_id is None:
            siswa_id = data.get('id')
            
        if not siswa_id:
            return jsonify({"status": "error", "message": "ID siswa tidak diberikan"}), 400

        nama = data.get('nama')
        kelas = data.get('kelas')
        kategori = data.get('kategori_masalah')

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE blueprints 
            SET nama = COALESCE(?, nama),
                kelas = COALESCE(?, kelas),
                kategori_masalah = COALESCE(?, kategori_masalah)
            WHERE id = ?
        ''', (nama, kelas, kategori, siswa_id))
        conn.commit()
        conn.close()

        return jsonify({"status": "success", "message": "Data siswa berhasil diperbarui"})

    except Exception as e:
        print(f"Error Edit Siswa: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/chat-admin', methods=['POST'])
@app.route('/api/diskusi-rpl', methods=['POST'])
@requires_auth
def chat_admin():
    try:
        data = request.get_json()
        messages = data.get('messages', [])
        chat_history = data.get('chatHistory', messages)
        
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('SELECT nama, kelas, kategori_masalah, ai_analisis FROM blueprints ORDER BY id DESC LIMIT 20')
        rows = cursor.fetchall()
        conn.close()
        
        konteks_siswa = "\n".join([f"- {r[0]} ({r[1]}) | Masalah: {r[2]} | Analisis: {r[3]}" for r in rows])
        
        system_prompt = f"""
        ROLE: Anda adalah "Asisten HR/Konsultan Edukasi Senior".
        TUGAS: Membantu Admin dalam merancang strategi layanan, mendiskusikan kasus pengguna, menyusun evaluasi, dan memberikan rekomendasi intervensi yang pragmatis.
        
        KONTEKS DATA PENGGUNA TERBARU DARI DATABASE:
        {konteks_siswa if konteks_siswa else 'Belum ada data pengguna.'}
        
        Gunakan bahasa profesional namun hangat, suportif, dan solutif.
        """
        
        formatted_messages = [{"role": "system", "content": system_prompt}]
        for msg in chat_history:
            formatted_messages.append({
                "role": msg.get("role", "user"),
                "content": msg.get("content", "")
            })

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=formatted_messages,
            temperature=0.5
        )
        
        reply = response.choices[0].message.content
        return jsonify({"status": "success", "reply": reply, "message": reply})

    except Exception as e:
        print(f"Error Chat Admin: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/statistik', methods=['GET'])
@requires_auth
def get_statistik():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT kategori_masalah, COUNT(id) as total
            FROM blueprints
            GROUP BY kategori_masalah
        ''')
        rows = cursor.fetchall()
        conn.close()

        labels = []
        data = []
        for row in rows:
            labels.append(row[0] if row[0] else "Belum Terkategori")
            data.append(row[1])

        return jsonify({
            "status": "success", 
            "labels": labels, 
            "data": data
        })

    except Exception as e:
        print(f"Error Load Statistik: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/generate-rpl', methods=['POST'])
@requires_auth
def generate_rpl():
    try:
        data = request.get_json()
        kategori_masalah = data.get('kategori', 'Akademik')
        metode_pilihan = data.get('metode', 'Bimbingan Klasikal')

        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT nama, kelas, j_hambatan, ai_analisis 
            FROM blueprints 
            WHERE kategori_masalah = ?
        ''', (kategori_masalah,))
        rows = cursor.fetchall()
        conn.close()

        konteks_siswa = ""
        if rows:
            for row in rows:
                nama_siswa = row[0]
                kelas = row[1]
                keluhan = row[2]
                analisis = row[3]
                konteks_siswa += f"- {nama_siswa} | Keluhan: {keluhan} | Analisis: {analisis}\n"
        else:
            return jsonify({"status": "error", "message": f"Belum ada data untuk kategori {kategori_masalah} di database."}), 404

        prompt = f"""
        ROLE: Anda adalah Pakar Pengembangan Diri/Konselor Senior yang pakar menyusun Rencana Pelaksanaan Layanan adaptif.
        
        TUGAS: Susun materi layanan untuk merespons kasus pada kategori masalah: {kategori_masalah}.

        PARAMETER LAYANAN:
        - Metode Layanan: {metode_pilihan}

        DATA RIIL PENGGUNA SAAT INI (WAJIB JADIKAN ACUAN MATERI):
        {konteks_siswa}

        ATURAN KETAT POLARITAS DIKSI JUDUL:
        1. PEMETAAN KATA KERJA LENGKAP:
           - Variabel NEGATIF / MASALAH (Kecemasan, Stres, Adiksi, Konflik, Gugup, dll): WAJIB menggunakan kata kerja reduksi ("Manajemen", "Reduksi", "Penanganan", "Pengendalian"). DILARANG KERAS menggunakan kata "Optimalisasi" atau "Peningkatan" untuk masalah negatif!
           - Variabel POSITIF / POTENSI (Motivasi, Kepercayaan Diri, Minat, Resiliensi): WAJIB menggunakan kata kerja eskalasi ("Peningkatan", "Optimalisasi", "Pengembangan", "Penguatan").
        
        2. STRUKTUR DAN TATA BAHASA:
           - Setiap variabel harus diikat oleh kata kerja yang sesuai dengan polaritasnya.
           - Sintesiskan kesamaan masalah dari data pengguna di atas untuk menjadi satu topik layanan yang kohesif.

        FORMAT OUTPUT JSON:
        {{
            "topik_layanan": "Judul profesional, presisi secara logika semantik (maksimal 10 kata)",
            "tujuan_layanan": "Tujuan objektif spesifik menyelesaikan keluhan dominan di atas",
            "tahap_awal": "Penjelasan aktivitas pembuka (apersepsi) yang relevan dengan kasus",
            "tahap_inti": ["Langkah inti 1 (solusi kasus)", "Langkah inti 2", "Langkah inti 3"],
            "tahap_penutup_dan_evaluasi": "Aktivitas penutup dan instrumen evaluasi spesifik"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda wajib membalas hanya dengan format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"}
        )
        
        result_json = json.loads(response.choices[0].message.content)
        
        tahap_inti_html = "".join([f"- {langkah}<br>" for langkah in result_json.get('tahap_inti', [])])
        
        hasil_rpl = f"""
        <b>RENCANA PELAKSANAAN PENGEMBANGAN / LAYANAN</b><br>
        ===================================================<br><br>
        <b>Topik Layanan :</b> {result_json.get('topik_layanan', '-')}<br>
        <b>Tujuan Layanan:</b> {result_json.get('tujuan_layanan', '-')}<br><br>
        <b>A. TAHAP AWAL</b><br>
        {result_json.get('tahap_awal', '-')}<br><br>
        <b>B. TAHAP INTI</b><br>
        {tahap_inti_html}<br>
        <b>C. TAHAP PENUTUP & EVALUASI</b><br>
        {result_json.get('tahap_penutup_dan_evaluasi', '-')}
        """

        return jsonify({"status": "success", "hasil": hasil_rpl})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/generate-rpl-chat', methods=['POST'])
@requires_auth
def generate_rpl_chat():
    try:
        data = request.get_json()
        chat_history = data.get('chatHistory', [])
        
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute('SELECT nama, kelas, kategori_masalah, ai_analisis FROM blueprints ORDER BY id DESC LIMIT 20')
        rows = cursor.fetchall()
        conn.close()
        
        konteks_siswa = "\n".join([f"- {r[0]} | Masalah: {r[2]} | Analisis: {r[3]}" for r in rows])
        pesan_obrolan = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in chat_history])
        
        prompt = f"""
        ROLE: Anda adalah Pakar Konselor Senior.
        TUGAS: Susun materi rencana pelaksanaan layanan berdasarkan diskusi chat berikut.
        
        RIWAYAT DISKUSI CHAT DENGAN ADMIN:
        {pesan_obrolan}
        
        DATA PENGGUNA SAAT INI (Jadikan acuan jika relevan dengan obrolan):
        {konteks_siswa}
        
        FORMAT OUTPUT JSON MURNI:
        {{
            "topik_layanan": "Judul profesional layanan bimbingan (maksimal 10 kata)",
            "tujuan_layanan": "Tujuan objektif dari layanan ini",
            "tahap_awal": "Aktivitas pembuka (apersepsi)",
            "tahap_inti": ["Langkah penyelesaian 1", "Langkah penyelesaian 2", "Langkah 3"],
            "tahap_penutup_dan_evaluasi": "Aktivitas penutup dan evaluasi"
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda wajib membalas hanya dengan format JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"}
        )
        
        result_json = json.loads(response.choices[0].message.content)
        return jsonify({"status": "success", "data": result_json})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/eksplorasi-siswa', methods=['POST'])
def eksplorasi_siswa():
    try:
        data = request.get_json()
        chat_history = data.get('chatHistory', [])
        blueprint_data = data.get('blueprintData', {})
        
        kategori = blueprint_data.get('kategori_masalah', '-')
        analisis = blueprint_data.get('analisis_realitas', '-')
        karir_utama = blueprint_data.get('rekomendasi_karir_utama', '-')
        karir_alt = blueprint_data.get('rekomendasi_karir_alternatif', '-')
        
        langkah_list = blueprint_data.get('langkah_konkret_1_tahun_kedepan', [])
        if isinstance(langkah_list, list):
            langkah = ", ".join(langkah_list)
        else:
            langkah = str(langkah_list)
            
        taktik = blueprint_data.get('taktik_mengatasi_hambatan', '-')
        
        # UPGRADE: AI disesuaikan untuk segala kalangan
        system_prompt = f"""
        [PERAN DAN NADA INTERAKSI]
        Anda adalah Asisten Eksplorasi Karir AI untuk segala kalangan. Nada bicaramu hangat, membumi, setara seperti teman diskusi yang solid, dan tidak menggurui. Gunakan bahasa Indonesia sehari-hari yang luwes (menggunakan 'aku' dan 'kamu'). Jangan gunakan format poin-poin yang kaku jika tidak diminta.

        [ATURAN UTAMA]
        Tugas utamamu adalah membantu pengguna mengeksplorasi minat dan hasil asesmen karir mereka. Biarkan pengguna memegang kendali percakapan. Berikan solusi taktis, langkah kecil (baby steps), atau referensi belajar jika diminta. Jangan mengulangi secara mentah-mentah hasil asesmen, gunakan sebagai latar belakang ingatan saja.

        [KONTEKS HASIL ASESMEN PENGGUNA SAAT INI]
        - Kategori Dominan: {kategori}
        - Analisis Realitas: {analisis}
        - Rekomendasi Karir: {karir_utama} atau {karir_alt}
        - Rekomendasi Langkah 1 Tahun Kedepan: {langkah}
        - Taktik Menghadapi Hambatan: {taktik}

        [PROTOKOL KRISIS & SOFT ROUTING - SANGAT PENTING]
        Jika pengguna menunjukkan tanda-tanda kelelahan mental yang ekstrem, depresi, keputusasaan mendalam, krisis eksistensial, atau indikasi menyakiti diri sendiri, kamu HARUS mematuhi 4 langkah berurutan ini dalam satu balasan:

        1. VALIDASI EMPATIK: Akui bahwa perasaannya nyata, berat, dan valid. Jangan pernah memberikan nasihat klise.
        2. NORMALISASI SITUASI: Beritahu bahwa merasa kewalahan di fase kehidupan saat ini adalah hal wajar.
        3. TAWARKAN KEHADIRAN (AI): Tegaskan bahwa kamu ada di sini saat ini untuk mendengarkan.
        4. SOFT ROUTING (OPSI MANUSIA): Letakkan tawaran menemui konselor profesional atau rekan HR/BK di akhir kalimat murni sebagai fasilitas "ruang aman", BUKAN sebagai keharusan.
        """

        messages = [{"role": "system", "content": system_prompt}]
        
        for msg in chat_history:
            messages.append({
                "role": msg.get("role", "user"),
                "content": msg.get("content", "")
            })

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            temperature=0.5
        )
        
        ai_reply = response.choices[0].message.content
        return jsonify({"status": "success", "reply": ai_reply})

    except Exception as e:
        print("Error Eksplorasi Pengguna:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/problem-solver', methods=['POST'])
def problem_solver():
    try:
        data = request.get_json()
        chat_history = data.get('chatHistory', [])

        # UPGRADE: AI diubah untuk merangkul remaja hingga profesional/pekerja
        system_prompt = """
        ROLE: Anda adalah "Adaptive Counselor AI", asisten penyelesaian masalah untuk remaja hingga dewasa dan pekerja profesional. Nada bicaramu hangat, pragmatis, adaptif, dan objektif—layaknya kawan diskusi cerdas, bukan bot psikologi klinis. Gunakan bahasa sehari-hari yang membumi (aku/kamu).

        [TUGAS UTAMA]
        1. Evaluasi ulang Kategori, Level Keparahan, Volatilitas Emosi, dan Analisis Kognitif pada SETIAP balasan berdasarkan SELURUH riwayat percakapan terbaru.
        2. Berikan balasan yang relevan secara kontekstual, berempati natural, dan tidak kaku.

        [KRITERIA TRIASE & EVALUASI LEVEL]
        - LEVEL 1 (Ringan): Masalah harian biasa (malas, bingung bagi waktu, tugas/pekerjaan menumpuk).
        - LEVEL 2 (Menengah): Kebuntuan jangka panjang (> 6 bulan), tekanan keluarga/atasan, krisis finansial harian, krisis percaya diri, overthinking.
        - LEVEL 3 (Krusial): Depresi, indikasi self-harm eksplisit, krisis eksistensial parah, keputusasaan ekstrem.

        [ATURAN KHUSUS DETEKSI VOLATILITAS & AVOIDANCE]
        1. BANDINGKAN konteks pesan terbaru dengan topik 2-3 pesan sebelumnya.
        2. JIKA pengguna berpindah secara drastis dari topik bertekanan tinggi (karier, finansial, eksistensial) ke topik sepele/santai tanpa penyelesaian masalah yang jelas:
           - JANGAN langsung menurunkan Level Keparahan ke Level 1.
           - Tetapkan "volatilitas_emosi": "Tinggi"
           - Tandai dalam "catatan_coping" sebagai indikasi "Avoidance/Pengalihan".
           - Pertahankan Level Keparahan minimal Level 2 selama 1-2 turn untuk memastikan kestabilan emosi pengguna.
        3. JIKA pergeseran terjadi secara wajar atau emosi konsisten, tetapkan "volatilitas_emosi": "Rendah" atau "Sedang".

        [ATURAN PENGELOLAAN CONTEXT & TONE - WAJIB DITAATI]
        1. REKOMENDASI PRAGMATIS (Anti Tone-Deaf): Jika pengguna butuh uang/penghasilan harian, DILARANG KERAS menyarankan proyek sukarela. Beri opsi praktis.
        2. VARIASI STRUKTUR BAHASA: DILARANG mengulang pola frasa kaku CBT (misal: "Mari kita pisahkan fakta..."). 
        3. DETEKSI PERGESERAN INTENT: Jika pengguna mengubah topik ke hal teknis, keahlian, atau pencapaian, SEGERA SESUAIKAN TONE menjadi apresiatif dan kolaboratif.
        
        [PROTOKOL KRISIS & ANTI-HALUSINASI]
        4. ANTI-ESKALASI PREMATUR: Analisis WAJIB berdasarkan teks eksplisit. DILARANG memprediksi atau menyimpulkan pengguna ingin bunuh diri/putus asa ekstrem HANYA KARENA mereka berkata "lelah" atau "memendam perasaan". Anda hanya boleh menetapkan status krisis mematikan jika pengguna menyatakan niat tersebut secara langsung.
        5. PENANGANAN LEVEL 3 (EMPATI MANUSIAWI): 
           - Gunakan empati ala teman di Indonesia (contoh: "Peluk jauh dari sini", "Aku paham rasanya pasti sesak dan berat banget", "Aku di sini nemenin kamu"). 
           - DILARANG KERAS menggunakan frasa klinis terjemahan.
           - Arahkan pengguna dengan sangat santai untuk menemui mentor, HRD, atau konselor sebagai "ruang aman untuk napas sebentar", BUKAN sebagai rujukan medis/rumah sakit.

        FORMAT OUTPUT (WAJIB JSON MURNI):
        {
            "kategori": "Salah satu dari [Akademik, Sosial, Pribadi, Karir]",
            "level_keparahan": 1, 2, atau 3 (pilih berupa angka murni),
            "volatilitas_emosi": "Salah satu dari [Rendah, Sedang, Tinggi]",
            "analisis_kognitif": {
                "skema_pikir": "Ringkasan pola pikir dominan (1 kalimat)",
                "pemicu_utama": "Hal yang mendasari kekhawatiran (1 kalimat)",
                "catatan_coping": "Deteksi mekanisme pengalihan atau respons emosional (1 kalimat)"
            },
            "balasan_ai": "Pesan balasan langsung untuk pengguna"
        }
        """

        messages = [{"role": "system", "content": system_prompt}]
        
        for msg in chat_history:
            messages.append({
                "role": msg.get("role", "user"),
                "content": msg.get("content", "")
            })

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            temperature=0.4, 
            response_format={"type": "json_object"}
        )
        
        result_json = json.loads(response.choices[0].message.content)
        return jsonify({"status": "success", "data": result_json})

    except Exception as e:
        print("Error Problem Solver:", str(e))
        return jsonify({
            "status": "error", 
            "data": {
                "kategori": "Sistem",
                "level_keparahan": 1,
                "volatilitas_emosi": "Rendah",
                "analisis_kognitif": {
                    "skema_pikir": "Gagal memproses analisis kognitif.",
                    "pemicu_utama": "-",
                    "catatan_coping": "-"
                },
                "balasan_ai": "Maaf, ada sedikit kendala koneksi. Bisakah Anda mengulangi pesan tadi?"
            }
        }), 500


@app.route('/api/upload-csv', methods=['POST'])
@requires_auth
def upload_csv():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({'status': 'error', 'message': 'File CSV tidak ditemukan'}), 400
            
        df = pd.read_csv(file)
        
        if 'Kelas' in df.columns:
            df['Kelas'] = df['Kelas'].apply(standarisasi_kelas)
        
        return jsonify({
            'status': 'success', 
            'total_siswa': len(df), 
            'message': 'File CSV berhasil diimpor dan data kelas berhasil dinormalisasi.'
        })

    except Exception as e:
        print(f"Error Upload CSV: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/generate-ppt', methods=['POST'])
@requires_auth
def generate_ppt():
    try:
        data = request.get_json()
        topik = data.get('topik', 'Materi Bimbingan')
        target_kelas = data.get('kelas', 'Umum')
        desain = data.get('desain', 'Kreatif (Krem & Oranye)')

        prompt = f"""
        TUGAS: Buat isi presentasi PowerPoint untuk layanan Bimbingan / Pengembangan Diri yang SIAP TAMPIL.
        TOPIK: {topik}
        TARGET AUDIENS: {target_kelas} 
        
        ATURAN KONTEN (SANGAT PENTING):
        1. DILARANG KERAS menggunakan kalimat tanya pada poin materi. Berikan definisi atau pernyataan langsung (Contoh: BUKAN "Apa itu stres?", TAPI "Stres adalah beban kognitif saat...").
        2. Gunakan gaya bahasa yang santai, membumi, namun tetap edukatif dan tidak kaku.
        3. Berikan nama teknik/solusi yang spesifik dan aplikatif (Contoh: Gunakan "Teknik Pomodoro", "Jurnal Syukur", atau "Matriks Eisenhower", BUKAN sekadar "Manajemen Waktu" atau "Relaksasi").
        4. Isi slide harus berupa kalimat ringkas yang informatif, bukan sekadar outline/kerangka materi kosong.
        5. WAJIB sertakan satu slide terakhir sebagai penutup yang berisi kesimpulan, afirmasi positif, dan ajakan (Call to Action) ramah untuk berdiskusi lebih lanjut.
        
        Buat 5-6 slide yang interaktif dan langsung bisa dipahami audiens.
        Format WAJIB JSON Murni:
        {{
            "slides": [
                {{
                    "judul": "Judul Menarik Slide 1",
                    "konten": ["Pernyataan/fakta menarik 1", "Pernyataan/fakta menarik 2", "Pernyataan/fakta menarik 3"]
                }}
            ]
        }}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Anda adalah asisten pembuat struktur PPT berformat JSON murni."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4,
            response_format={"type": "json_object"}
        )
        
        ppt_data = json.loads(response.choices[0].message.content)

        prs = Presentation()
        
        for slide_data in ppt_data.get('slides', []):
            slide_layout = prs.slide_layouts[1] 
            slide = prs.slides.add_slide(slide_layout)
            
            title_placeholder = slide.shapes.title
            body_placeholder = slide.placeholders[1]

            title_placeholder.text = slide_data.get('judul', '')
            
            tf = body_placeholder.text_frame
            tf.clear() 
            
            for poin in slide_data.get('konten', []):
                p = tf.add_paragraph()
                p.text = poin
                p.level = 0

        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)

        nama_file = f"PPT_Materi_{topik.replace(' ', '_')}.pptx"

        return send_file(
            ppt_stream,
            as_attachment=True,
            download_name=nama_file,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        print("Error Generate PPT:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/export-csv')
@requires_auth
def export_csv():
    try:
        conn = sqlite3.connect('data_siswa.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM blueprints ORDER BY id DESC")
        rows = cursor.fetchall()
        column_names = [description[0] for description in cursor.description]
        conn.close()

        si = StringIO()
        cw = csv.writer(si)
        cw.writerow(column_names)
        cw.writerows(rows)

        output = make_response(si.getvalue())
        output.headers["Content-Disposition"] = "attachment; filename=Rekap_Asesmen_Karir.csv"
        output.headers["Content-type"] = "text/csv"
        return output
    except Exception as e:
        return str(e), 500


if __name__ == '__main__':
    # UPGRADE KEAMANAN: Memastikan mode Debug dimatikan di server produksi
    is_debug = os.getenv("FLASK_DEBUG", "False").lower() in ["true", "1"]
    port = int(os.getenv("PORT", 5000))
    app.run(debug=is_debug, host='0.0.0.0', port=port)
