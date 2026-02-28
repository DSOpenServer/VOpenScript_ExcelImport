<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PHP 8 Excel Import · Stok Aktarım</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Syne:wght@400;600;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg:        #0d0f12;
            --surface:   #13161b;
            --surface2:  #1c2028;
            --border:    #262b35;
            --accent:    #3b82f6;
            --accent2:   #10b981;
            --warn:      #f59e0b;
            --danger:    #ef4444;
            --text:      #e2e8f0;
            --muted:     #64748b;
            --code-bg:   #0a0c10;
        }

        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        html { scroll-behavior: smooth; }

        body {
            background: var(--bg);
            color: var(--text);
            font-family: 'DM Sans', sans-serif;
            font-size: 15px;
            line-height: 1.7;
            min-height: 100vh;
        }

        /* ── NOISE TEXTURE OVERLAY ── */
        body::before {
            content: '';
            position: fixed;
            inset: 0;
            background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.04'/%3E%3C/svg%3E");
            pointer-events: none;
            z-index: 0;
        }

        /* ── LAYOUT ── */
        .wrapper {
            max-width: 900px;
            margin: 0 auto;
            padding: 0 24px;
            position: relative;
            z-index: 1;
        }

        /* ── HERO ── */
        .hero {
            padding: 80px 0 60px;
            border-bottom: 1px solid var(--border);
        }

        .badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(59,130,246,.1);
            border: 1px solid rgba(59,130,246,.25);
            color: var(--accent);
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            padding: 4px 10px;
            border-radius: 20px;
            margin-bottom: 24px;
            letter-spacing: .05em;
        }

        .badge::before { content: '●'; font-size: 8px; }

        h1 {
            font-family: 'Syne', sans-serif;
            font-size: clamp(2rem, 5vw, 3.2rem);
            font-weight: 800;
            line-height: 1.1;
            letter-spacing: -.02em;
            margin-bottom: 16px;
        }

        h1 span {
            background: linear-gradient(135deg, var(--accent), var(--accent2));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        .hero-desc {
            color: var(--muted);
            font-size: 16px;
            font-weight: 300;
            max-width: 560px;
            margin-bottom: 32px;
        }

        .hero-badges {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }

        .pill {
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            padding: 3px 10px;
            border-radius: 4px;
            border: 1px solid var(--border);
            color: var(--muted);
        }

        .pill.green  { color: var(--accent2); border-color: rgba(16,185,129,.3); background: rgba(16,185,129,.05); }
        .pill.blue   { color: var(--accent);  border-color: rgba(59,130,246,.3); background: rgba(59,130,246,.05); }
        .pill.yellow { color: var(--warn);    border-color: rgba(245,158,11,.3); background: rgba(245,158,11,.05); }

        /* ── NAV ── */
        nav {
            position: sticky;
            top: 0;
            background: rgba(13,15,18,.85);
            backdrop-filter: blur(12px);
            border-bottom: 1px solid var(--border);
            z-index: 100;
            padding: 0;
        }

        nav .wrapper {
            display: flex;
            gap: 0;
            overflow-x: auto;
        }

        nav a {
            display: block;
            padding: 14px 18px;
            color: var(--muted);
            text-decoration: none;
            font-size: 13px;
            font-weight: 500;
            white-space: nowrap;
            border-bottom: 2px solid transparent;
            transition: color .2s, border-color .2s;
        }

        nav a:hover { color: var(--text); border-color: var(--border); }
        nav a.active { color: var(--accent); border-color: var(--accent); }

        /* ── SECTIONS ── */
        section {
            padding: 64px 0;
            border-bottom: 1px solid var(--border);
        }

        section:last-child { border-bottom: none; }

        .section-label {
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            color: var(--accent);
            letter-spacing: .1em;
            text-transform: uppercase;
            margin-bottom: 8px;
        }

        h2 {
            font-family: 'Syne', sans-serif;
            font-size: 1.6rem;
            font-weight: 700;
            letter-spacing: -.01em;
            margin-bottom: 24px;
        }

        h3 {
            font-family: 'Syne', sans-serif;
            font-size: 1.1rem;
            font-weight: 600;
            margin-bottom: 12px;
            margin-top: 32px;
        }

        p { color: #94a3b8; margin-bottom: 16px; }

        /* ── FILE TREE ── */
        .file-tree {
            background: var(--code-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 20px 24px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 13px;
            line-height: 2;
        }

        .file-tree .dir  { color: var(--accent); }
        .file-tree .file { color: var(--text); }
        .file-tree .cmt  { color: var(--muted); }
        .file-tree .new  { color: var(--accent2); }
        .file-tree .ind  { color: var(--border); margin-right: 4px; }

        /* ── CODE BLOCKS ── */
        .code-wrap {
            position: relative;
            margin: 20px 0;
        }

        .code-label {
            position: absolute;
            top: -1px;
            left: 16px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 10px;
            color: var(--muted);
            background: var(--code-bg);
            border: 1px solid var(--border);
            border-top: none;
            padding: 2px 8px;
            border-radius: 0 0 4px 4px;
            letter-spacing: .05em;
        }

        pre {
            background: var(--code-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 24px 20px 20px;
            overflow-x: auto;
            font-family: 'JetBrains Mono', monospace;
            font-size: 13px;
            line-height: 1.8;
            tab-size: 4;
        }

        /* Syntax colors */
        .kw  { color: #c084fc; }   /* keyword */
        .fn  { color: #60a5fa; }   /* function */
        .st  { color: #34d399; }   /* string */
        .cm  { color: #475569; font-style: italic; } /* comment */
        .va  { color: #f8b4d9; }   /* variable */
        .nu  { color: #fbbf24; }   /* number/const */
        .cl  { color: #fb923c; }   /* class */
        .hl  { background: rgba(59,130,246,.12); display: block; margin: 0 -20px; padding: 0 20px; border-left: 2px solid var(--accent); }

        /* ── CARDS ── */
        .card-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
            gap: 16px;
            margin-top: 24px;
        }

        .card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 20px;
            transition: border-color .2s, transform .2s;
        }

        .card:hover {
            border-color: var(--accent);
            transform: translateY(-2px);
        }

        .card-icon {
            font-size: 22px;
            margin-bottom: 10px;
        }

        .card h4 {
            font-family: 'Syne', sans-serif;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 6px;
        }

        .card p {
            font-size: 13px;
            color: var(--muted);
            margin: 0;
        }

        /* ── URL TABLE ── */
        .url-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            margin-top: 20px;
        }

        .url-table th {
            text-align: left;
            padding: 10px 14px;
            background: var(--surface2);
            color: var(--muted);
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            letter-spacing: .05em;
            border-bottom: 1px solid var(--border);
        }

        .url-table td {
            padding: 12px 14px;
            border-bottom: 1px solid var(--border);
            vertical-align: top;
        }

        .url-table tr:last-child td { border-bottom: none; }

        .url-table tr:hover td { background: var(--surface); }

        .url-table .url-cell {
            font-family: 'JetBrains Mono', monospace;
            color: var(--accent);
            font-size: 12px;
        }

        .method {
            display: inline-block;
            font-family: 'JetBrains Mono', monospace;
            font-size: 10px;
            padding: 2px 7px;
            border-radius: 3px;
            font-weight: 600;
        }

        .method.get  { background: rgba(16,185,129,.15); color: var(--accent2); border: 1px solid rgba(16,185,129,.3); }
        .method.post { background: rgba(59,130,246,.15); color: var(--accent);  border: 1px solid rgba(59,130,246,.3); }

        /* ── ALERT BOXES ── */
        .alert {
            display: flex;
            gap: 12px;
            padding: 14px 16px;
            border-radius: 8px;
            margin: 20px 0;
            font-size: 14px;
        }

        .alert-icon { font-size: 16px; flex-shrink: 0; margin-top: 1px; }
        .alert p { margin: 0; color: inherit; font-size: 14px; }

        .alert.info    { background: rgba(59,130,246,.08);  border: 1px solid rgba(59,130,246,.25);  color: #93c5fd; }
        .alert.warn    { background: rgba(245,158,11,.08);  border: 1px solid rgba(245,158,11,.25);  color: #fcd34d; }
        .alert.success { background: rgba(16,185,129,.08);  border: 1px solid rgba(16,185,129,.25);  color: #6ee7b7; }
        .alert.danger  { background: rgba(239,68,68,.08);   border: 1px solid rgba(239,68,68,.25);   color: #fca5a5; }

        /* ── STEPS ── */
        .steps { counter-reset: step; list-style: none; }

        .steps li {
            counter-increment: step;
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
        }

        .steps li::before {
            content: counter(step);
            flex-shrink: 0;
            width: 28px;
            height: 28px;
            background: var(--surface2);
            border: 1px solid var(--border);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: 'JetBrains Mono', monospace;
            font-size: 12px;
            color: var(--accent);
            margin-top: 2px;
        }

        .steps li > div h4 {
            font-family: 'Syne', sans-serif;
            font-weight: 600;
            margin-bottom: 4px;
        }

        .steps li > div p { font-size: 14px; margin: 0; }

        /* ── INLINE CODE ── */
        code {
            font-family: 'JetBrains Mono', monospace;
            font-size: 12px;
            background: var(--surface2);
            border: 1px solid var(--border);
            padding: 1px 6px;
            border-radius: 3px;
            color: var(--accent2);
        }

        /* ── FOOTER ── */
        footer {
            padding: 40px 0;
            text-align: center;
            color: var(--muted);
            font-size: 13px;
        }

        footer a { color: var(--accent); text-decoration: none; }
        footer a:hover { text-decoration: underline; }

        /* ── SCROLLBAR ── */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: var(--bg); }
        ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

        /* ── ANIMATIONS ── */
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(20px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        .hero > * {
            animation: fadeUp .5s ease both;
        }

        .hero > *:nth-child(1) { animation-delay: .05s; }
        .hero > *:nth-child(2) { animation-delay: .12s; }
        .hero > *:nth-child(3) { animation-delay: .18s; }
        .hero > *:nth-child(4) { animation-delay: .24s; }
    </style>
</head>
<body>

<!-- ══════════════════════════════════════════
     HERO
════════════════════════════════════════════ -->
<div class="wrapper">
    <div class="hero">
        <div class="badge">PHP 8 · PDO · Excel Import</div>
        <h1>Stok Aktarım<br><span>Excel → MySQL</span></h1>
        <p class="hero-desc">
            PHP 8 uyumlu, PDO prepared statement kullanan, XSS ve SQL injection korumalı
            Excel dosyasından stok verisi aktarım scripti.
        </p>
        <div class="hero-badges">
            <span class="pill blue">PHP 8.x</span>
            <span class="pill green">PDO / Prepared Statements</span>
            <span class="pill green">XSS Koruması</span>
            <span class="pill yellow">BIFF7 / BIFF8 / XLSX</span>
            <span class="pill">MIT Lisans</span>
        </div>
    </div>
</div>

<!-- ══════════════════════════════════════════
     NAV
════════════════════════════════════════════ -->
<nav>
    <div class="wrapper">
        <a href="#genel-bakis" class="active">Genel Bakış</a>
        <a href="#kurulum">Kurulum</a>
        <a href="#dosya-yapisi">Dosya Yapısı</a>
        <a href="#url-referans">URL Referans</a>
        <a href="#kod-ornekleri">Kod Örnekleri</a>
        <a href="#db-class">dbClass</a>
        <a href="#excel-reader">Excel Reader</a>
        <a href="#guvenlik">Güvenlik</a>
    </div>
</nav>

<!-- ══════════════════════════════════════════
     GENEL BAKIŞ
════════════════════════════════════════════ -->
<div class="wrapper">
<section id="genel-bakis">
    <div class="section-label">// 01 · GENEL BAKIŞ</div>
    <h2>Proje Hakkında</h2>
    <p>
        Bu proje, <strong>PHP 8</strong> ile yazılmış, <code>.xls</code> ve <code>.xlsx</code> formatındaki
        Excel dosyalarından MySQL veritabanına stok verisi aktaran bir import sistemidir.
        Orijinal PHP 4/5 kodu, PHP 8 standartlarına uygun şekilde yeniden yazılmıştır.
    </p>

    <div class="card-grid">
        <div class="card">
            <div class="card-icon">⚡</div>
            <h4>PHP 8 Native</h4>
            <p>Typed properties, union types, named arguments ve null-safe operatör kullanımı.</p>
        </div>
        <div class="card">
            <div class="card-icon">🔒</div>
            <h4>Güvenli Sorgular</h4>
            <p>Tüm veritabanı işlemleri PDO Prepared Statements ile yapılır. SQL injection riski sıfır.</p>
        </div>
        <div class="card">
            <div class="card-icon">📊</div>
            <h4>Çoklu Format</h4>
            <p>XLS (BIFF7/BIFF8), XLSX ve CSV formatlarını destekler. Otomatik format tespiti.</p>
        </div>
        <div class="card">
            <div class="card-icon">🛡️</div>
            <h4>XSS Koruması</h4>
            <p>Tüm çıktılar <code>htmlspecialchars()</code> ile temizlenir.</p>
        </div>
    </div>
</section>

<!-- ══════════════════════════════════════════
     KURULUM
════════════════════════════════════════════ -->
<section id="kurulum">
    <div class="section-label">// 02 · KURULUM</div>
    <h2>Hızlı Kurulum</h2>

    <ol class="steps">
        <li>
            <div>
                <h4>Repoyu klonla</h4>
                <p>GitHub'dan projeyi yerel ortamınıza indirin.</p>
                <div class="code-wrap" style="margin-top:10px">
                    <pre><span class="cm"># XAMPP kullanıyorsanız htdocs altına</span>
git clone https://github.com/kullanici/php8-excel-import.git
<span class="kw">cd</span> php8-excel-import</pre>
                </div>
            </div>
        </li>
        <li>
            <div>
                <h4>Veritabanını oluştur</h4>
                <p>Aşağıdaki SQL'i phpMyAdmin veya MySQL CLI ile çalıştırın.</p>
                <div class="code-wrap" style="margin-top:10px">
                    <span class="code-label">SQL</span>
                    <pre><span class="kw">CREATE DATABASE</span> voidsysweb <span class="kw">CHARACTER SET</span> utf8mb4 <span class="kw">COLLATE</span> utf8mb4_unicode_ci;

<span class="kw">USE</span> voidsysweb;

<span class="kw">CREATE TABLE</span> <span class="cl">stok</span> (
  <span class="va">id</span>          <span class="nu">INT</span> <span class="kw">AUTO_INCREMENT PRIMARY KEY</span>,
  <span class="va">stokkod</span>     <span class="nu">VARCHAR</span>(<span class="nu">50</span>)  <span class="kw">NOT NULL UNIQUE</span>,
  <span class="va">stokad</span>      <span class="nu">VARCHAR</span>(<span class="nu">255</span>) <span class="kw">NOT NULL</span>,
  <span class="va">grupkod</span>     <span class="nu">VARCHAR</span>(<span class="nu">50</span>)  <span class="kw">DEFAULT</span> <span class="st">''</span>,
  <span class="va">ekgrupkod</span>   <span class="nu">VARCHAR</span>(<span class="nu">50</span>)  <span class="kw">DEFAULT</span> <span class="st">''</span>,
  <span class="va">birim</span>       <span class="nu">VARCHAR</span>(<span class="nu">20</span>)  <span class="kw">DEFAULT</span> <span class="st">''</span>,
  <span class="va">kdvoran</span>     <span class="nu">DECIMAL</span>(<span class="nu">5,2</span>) <span class="kw">DEFAULT</span> <span class="nu">0</span>,
  <span class="va">aciklama</span>    <span class="nu">TEXT</span>,
  <span class="va">alisfiyat</span>   <span class="nu">DECIMAL</span>(<span class="nu">10,2</span>) <span class="kw">DEFAULT</span> <span class="nu">0</span>,
  <span class="va">satisfiyat</span>  <span class="nu">DECIMAL</span>(<span class="nu">10,2</span>) <span class="kw">DEFAULT</span> <span class="nu">0</span>,
  <span class="va">created_at</span>  <span class="nu">TIMESTAMP</span> <span class="kw">DEFAULT</span> <span class="nu">CURRENT_TIMESTAMP</span>
) <span class="nu">ENGINE</span>=InnoDB;</pre>
                </div>
            </div>
        </li>
        <li>
            <div>
                <h4>Veritabanı bağlantısını yapılandır</h4>
                <p><code>classes/dbClass.php</code> dosyasındaki bilgileri düzenleyin.</p>
                <div class="code-wrap" style="margin-top:10px">
                    <span class="code-label">classes/dbClass.php</span>
                    <pre><span class="kw">private</span> <span class="va">$host</span>    = <span class="st">"localhost"</span>;
<span class="kw">private</span> <span class="va">$user</span>    = <span class="st">"root"</span>;       <span class="cm">// ← MySQL kullanıcı adı</span>
<span class="kw">private</span> <span class="va">$pass</span>    = <span class="st">""</span>;          <span class="cm">// ← MySQL şifresi</span>
<span class="kw">private</span> <span class="va">$db</span>      = <span class="st">"voidsysweb"</span>; <span class="cm">// ← Veritabanı adı</span>
<span class="kw">private</span> <span class="va">$charset</span> = <span class="st">"utf8mb4"</span>;</pre>
                </div>
            </div>
        </li>
        <li>
            <div>
                <h4>uploads/ klasörünü oluştur</h4>
                <p>Geçici dosyaların yazılabilmesi için klasör ve izinleri ayarlayın.</p>
                <div class="code-wrap" style="margin-top:10px">
                    <pre>mkdir excelimport/uploads
chmod 755 excelimport/uploads   <span class="cm"># Linux/macOS</span></pre>
                </div>
            </div>
        </li>
        <li>
            <div>
                <h4>Tarayıcıdan aç</h4>
                <p>XAMPP çalışıyorsa doğrudan erişin.</p>
                <div class="code-wrap" style="margin-top:10px">
                    <pre>http://localhost/php8-excel-import/excelimport/stokimport.php</pre>
                </div>
            </div>
        </li>
    </ol>
</section>

<!-- ══════════════════════════════════════════
     DOSYA YAPISI
════════════════════════════════════════════ -->
<section id="dosya-yapisi">
    <div class="section-label">// 03 · DOSYA YAPISI</div>
    <h2>Proje Klasör Yapısı</h2>

    <div class="file-tree">
<span class="dir">php8-excel-import/</span><br>
<span class="ind">├──</span> <span class="new">index.php</span>                   <span class="cmt">&lt;!-- Bu dokümantasyon sayfası --&gt;</span><br>
<span class="ind">├──</span> <span class="file">classes_include.php</span>         <span class="cmt">&lt;!-- Sınıf yükleyici --&gt;</span><br>
<span class="ind">├──</span> <span class="dir">classes/</span><br>
<span class="ind">│   └──</span> <span class="new">dbClass.php</span>             <span class="cmt">&lt;!-- PDO veritabanı sınıfı --&gt;</span><br>
<span class="ind">├──</span> <span class="dir">excelimport/</span><br>
<span class="ind">│   ├──</span> <span class="new">stokimport.php</span>          <span class="cmt">&lt;!-- Ana aktarım sayfası --&gt;</span><br>
<span class="ind">│   ├──</span> <span class="file">stoklar.xlsx</span>            <span class="cmt">&lt;!-- Örnek Excel şablonu --&gt;</span><br>
<span class="ind">│   ├──</span> <span class="dir">uploads/</span>                <span class="cmt">&lt;!-- Geçici yükleme klasörü (yazılabilir) --&gt;</span><br>
<span class="ind">│   └──</span> <span class="dir">vendor/</span><br>
<span class="ind">│       ├──</span> <span class="new">SpreadsheetReader.php</span>   <span class="cmt">&lt;!-- PHP 8 uyumlu reader --&gt;</span><br>
<span class="ind">│       ├──</span> <span class="new">SpreadsheetReader_XLS.php</span><br>
<span class="ind">│       ├──</span> <span class="new">SpreadsheetReader_XLSX.php</span><br>
<span class="ind">│       ├──</span> <span class="new">SpreadsheetReader_CSV.php</span><br>
<span class="ind">│       ├──</span> <span class="new">SpreadsheetReader_ODS.php</span><br>
<span class="ind">│       └──</span> <span class="new">excel_reader2.php</span>       <span class="cmt">&lt;!-- BIFF7/BIFF8 okuyucu --&gt;</span><br>
<span class="ind">└──</span> <span class="file">README.md</span>
    </div>

    <div class="alert info" style="margin-top:20px">
        <span class="alert-icon">ℹ️</span>
        <p><strong>Yeşil dosyalar</strong> PHP 8'e dönüştürülmüş versiyonlardır. Orijinal PHP 4/5 kodlarının yerini alır.</p>
    </div>
</section>

<!-- ══════════════════════════════════════════
     URL REFERANS
════════════════════════════════════════════ -->
<section id="url-referans">
    <div class="section-label">// 04 · URL REFERANS</div>
    <h2>URL Referansı</h2>
    <p>Projedeki tüm sayfalara ait URL listesi. <code>localhost</code> yerine kendi domain adınızı yazın.</p>

    <table class="url-table">
        <thead>
            <tr>
                <th>METOD</th>
                <th>URL</th>
                <th>AÇIKLAMA</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><span class="method get">GET</span></td>
                <td class="url-cell">http://localhost/php8-excel-import/</td>
                <td>Dokümantasyon ana sayfası (bu sayfa)</td>
            </tr>
            <tr>
                <td><span class="method get">GET</span></td>
                <td class="url-cell">http://localhost/php8-excel-import/excelimport/stokimport.php</td>
                <td>Stok aktarım formu — Excel yükleme ekranı</td>
            </tr>
            <tr>
                <td><span class="method post">POST</span></td>
                <td class="url-cell">http://localhost/php8-excel-import/excelimport/stokimport.php</td>
                <td>Excel dosyasını işle ve veritabanına kaydet</td>
            </tr>
            <tr>
                <td><span class="method get">GET</span></td>
                <td class="url-cell">http://localhost/php8-excel-import/excelimport/stoklar.xlsx</td>
                <td>Örnek Excel şablonunu indir</td>
            </tr>
        </tbody>
    </table>

    <h3>POST İstek Parametreleri</h3>
    <table class="url-table">
        <thead>
            <tr>
                <th>PARAMETRE</th>
                <th>TİP</th>
                <th>AÇIKLAMA</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><code>file</code></td>
                <td><span class="pill blue">FILE</span></td>
                <td>Yüklenecek <code>.xls</code> veya <code>.xlsx</code> dosyası</td>
            </tr>
            <tr>
                <td><code>import</code></td>
                <td><span class="pill">STRING</span></td>
                <td>Submit butonunun <code>name</code> değeri — formun gönderildiğini belirtir</td>
            </tr>
        </tbody>
    </table>

    <h3>Excel Sütun Sıralaması</h3>
    <table class="url-table">
        <thead>
            <tr>
                <th>SÜTUN</th>
                <th>INDEX</th>
                <th>ALAN</th>
                <th>ÖRNEK</th>
            </tr>
        </thead>
        <tbody>
            <tr><td>A</td><td><code>$Row[0]</code></td><td>stokkod</td><td>STK001</td></tr>
            <tr><td>B</td><td><code>$Row[1]</code></td><td>stokad</td><td>Ahşap Masa</td></tr>
            <tr><td>C</td><td><code>$Row[2]</code></td><td>grupkod</td><td>MOB</td></tr>
            <tr><td>D</td><td><code>$Row[3]</code></td><td>ekgrupkod</td><td>MOB-AH</td></tr>
            <tr><td>E</td><td><code>$Row[4]</code></td><td>birim</td><td>ADET</td></tr>
            <tr><td>F</td><td><code>$Row[5]</code></td><td>kdvoran</td><td>18</td></tr>
            <tr><td>G</td><td><code>$Row[6]</code></td><td>aciklama</td><td>180x80 cm doğal ahşap</td></tr>
            <tr><td>H</td><td><code>$Row[7]</code></td><td>alisfiyat</td><td>1200.00</td></tr>
            <tr><td>I</td><td><code>$Row[8]</code></td><td>satisfiyat</td><td>1850.00</td></tr>
        </tbody>
    </table>

    <div class="alert warn">
        <span class="alert-icon">⚠️</span>
        <p><strong>İlk satır başlık satırıdır</strong> — script otomatik olarak <code>$satir === 0</code> kontrolüyle atlar. Excel'in 1. satırına sütun adlarını yazın, veriler 2. satırdan başlasın.</p>
    </div>
</section>

<!-- ══════════════════════════════════════════
     KOD ÖRNEKLERİ
════════════════════════════════════════════ -->
<section id="kod-ornekleri">
    <div class="section-label">// 05 · KOD ÖRNEKLERİ</div>
    <h2>Kod Örnekleri</h2>

    <h3>Excel Dosyasını Okuma ve Satır İşleme</h3>
    <p>SpreadsheetReader ile çoklu sayfa desteğiyle Excel verilerini nasıl okuyacağınız:</p>

    <div class="code-wrap">
        <span class="code-label">stokimport.php — Excel okuma döngüsü</span>
        <pre><span class="va">$Reader</span>     = <span class="kw">new</span> <span class="cl">SpreadsheetReader</span>(<span class="va">$targetPath</span>);
<span class="va">$sheetCount</span> = <span class="fn">count</span>(<span class="va">$Reader</span>-><span class="fn">sheets</span>());

<span class="kw">for</span> (<span class="va">$i</span> = <span class="nu">0</span>; <span class="va">$i</span> < <span class="va">$sheetCount</span>; <span class="va">$i</span>++) {
    <span class="va">$Reader</span>-><span class="fn">ChangeSheet</span>(<span class="va">$i</span>);
    <span class="va">$satir</span> = <span class="nu">0</span>;

    <span class="kw">foreach</span> (<span class="va">$Reader</span> <span class="kw">as</span> <span class="va">$Row</span>) {
        <span class="cm">// İlk satır (başlık) atlanır</span>
        <span class="kw">if</span> (<span class="va">$satir</span> === <span class="nu">0</span>) { <span class="va">$satir</span>++; <span class="kw">continue</span>; }

        <span class="cm">// Null-safe veri okuma</span>
        <span class="hl">        <span class="va">$stokkod</span> = <span class="fn">trim</span>((<span class="kw">string</span>)(<span class="va">$Row</span>[<span class="nu">0</span>] ?? <span class="st">''</span>));</span>
        <span class="va">$stokad</span>  = <span class="fn">trim</span>((<span class="kw">string</span>)(<span class="va">$Row</span>[<span class="nu">1</span>] ?? <span class="st">''</span>));
        <span class="cm">// ... diğer alanlar</span>

        <span class="va">$satir</span>++;
    }
}</pre>
    </div>

    <h3>Prepared Statement ile Veri Ekleme</h3>
    <p>SQL injection'a karşı tam koruma sağlayan sorgular:</p>

    <div class="code-wrap">
        <span class="code-label">stokimport.php — Güvenli INSERT</span>
        <pre><span class="cm">// ✅ DOĞRU — Prepared Statement</span>
<span class="hl">        <span class="va">$dba</span>-><span class="fn">safeQuery</span>(</span>
<span class="hl">            <span class="st">"INSERT INTO stok</span></span>
<span class="hl">             <span class="st">(stokkod, stokad, grupkod, ekgrupkod, birim, kdvoran, aciklama, alisfiyat, satisfiyat)</span></span>
<span class="hl">             <span class="st">VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"</span>,</span>
<span class="hl">            [<span class="va">$stokkod</span>, <span class="va">$stokad</span>, <span class="va">$grupkod</span>, <span class="va">$ekgrupkod</span>,</span>
<span class="hl">             <span class="va">$birim</span>, <span class="va">$kdvoran</span>, <span class="va">$aciklama</span>, <span class="va">$alisfiyat</span>, <span class="va">$satisfiyat</span>]</span>
<span class="hl">        );</span>

<span class="cm">// ❌ YANLIŞ — SQL Injection açığı (eski yöntem)</span>
<span class="va">$dba</span>-><span class="fn">query</span>(<span class="st">"INSERT INTO stok VALUES('</span><span class="va">$stokkod</span><span class="st">',...)"</span>);</pre>
    </div>

    <h3>Mükerrer Kayıt Kontrolü</h3>

    <div class="code-wrap">
        <span class="code-label">stokimport.php — Duplicate check</span>
        <pre><span class="va">$kontrolStmt</span> = <span class="va">$dba</span>-><span class="fn">safeQuery</span>(
    <span class="st">"SELECT COUNT(id) AS sayi FROM stok WHERE stokkod = ?"</span>,
    [<span class="va">$stokkod</span>]
);
<span class="va">$sonuc</span> = <span class="va">$dba</span>-><span class="fn">fetch_object</span>(<span class="va">$kontrolStmt</span>);

<span class="kw">if</span> ((<span class="kw">int</span>)<span class="va">$sonuc</span>->sayi > <span class="nu">0</span>) {
    <span class="va">$message</span> .= <span class="fn">htmlspecialchars</span>(<span class="va">$stokkod</span>, <span class="nu">ENT_QUOTES</span>, <span class="st">'UTF-8'</span>)
        . <span class="st">' zaten var, atlandı.&lt;br&gt;'</span>;
}</pre>
    </div>

    <h3>Güvenli HTML Çıktısı</h3>

    <div class="code-wrap">
        <span class="code-label">stokimport.php — XSS korumalı çıktı</span>
        <pre><span class="cm">// ✅ DOĞRU</span>
<span class="hl">        <span class="kw">&lt;?=</span> <span class="fn">htmlspecialchars</span>((<span class="kw">string</span>)<span class="va">$stok</span>->stokkod, <span class="nu">ENT_QUOTES</span>, <span class="st">'UTF-8'</span>) <span class="kw">?&gt;</span></span>

<span class="cm">// ❌ YANLIŞ — Direkt echo (XSS açığı)</span>
<span class="kw">&lt;?php</span> <span class="kw">echo</span> <span class="va">$sonuc</span>->stokkod; <span class="kw">?&gt;</span></pre>
    </div>
</section>

<!-- ══════════════════════════════════════════
     DB CLASS
════════════════════════════════════════════ -->
<section id="db-class">
    <div class="section-label">// 06 · DBCLASS</div>
    <h2>dbClass Referansı</h2>
    <p>PDO tabanlı veritabanı sınıfı. Tüm sorgular bu sınıf üzerinden yapılır.</p>

    <table class="url-table">
        <thead>
            <tr>
                <th>METOD</th>
                <th>PARAMETRELER</th>
                <th>DÖNDÜRÜR</th>
                <th>AÇIKLAMA</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><code>connect()</code></td>
                <td>—</td>
                <td><code>PDO</code></td>
                <td>Veritabanı bağlantısını kurar</td>
            </tr>
            <tr>
                <td><code>safeQuery($sql, $params)</code></td>
                <td><code>string, array</code></td>
                <td><code>PDOStatement</code></td>
                <td>Prepared statement ile güvenli sorgu çalıştırır</td>
            </tr>
            <tr>
                <td><code>query($sql)</code></td>
                <td><code>string</code></td>
                <td><code>PDOStatement</code></td>
                <td>Parametresiz basit sorgular için (sadece SELECT)</td>
            </tr>
            <tr>
                <td><code>fetch_object($stmt)</code></td>
                <td><code>PDOStatement</code></td>
                <td><code>object|false</code></td>
                <td>Sonuçtan bir satırı nesne olarak döndürür</td>
            </tr>
            <tr>
                <td><code>num_rows($stmt)</code></td>
                <td><code>PDOStatement</code></td>
                <td><code>int</code></td>
                <td>Etkilenen / dönen satır sayısı</td>
            </tr>
            <tr>
                <td><code>insert_id()</code></td>
                <td>—</td>
                <td><code>string</code></td>
                <td>Son INSERT işleminin ID'sini döndürür</td>
            </tr>
        </tbody>
    </table>

    <div class="code-wrap" style="margin-top:24px">
        <span class="code-label">classes/dbClass.php — Tam kaynak</span>
        <pre><span class="kw">class</span> <span class="cl">dbClass</span> {
    <span class="kw">private</span> <span class="kw">string</span> <span class="va">$host</span>    = <span class="st">"localhost"</span>;
    <span class="kw">private</span> <span class="kw">string</span> <span class="va">$user</span>    = <span class="st">"root"</span>;
    <span class="kw">private</span> <span class="kw">string</span> <span class="va">$pass</span>    = <span class="st">""</span>;
    <span class="kw">private</span> <span class="kw">string</span> <span class="va">$db</span>      = <span class="st">"voidsysweb"</span>;
    <span class="kw">private</span> <span class="kw">string</span> <span class="va">$charset</span> = <span class="st">"utf8mb4"</span>;
    <span class="kw">public</span>  <span class="cl">PDO</span>    <span class="va">$pdo</span>;

    <span class="kw">public function</span> <span class="fn">connect</span>(): <span class="cl">PDO</span> {
        <span class="va">$dsn</span> = <span class="st">"mysql:host=<span class="va">$this</span>->host;dbname=<span class="va">$this</span>->db;charset=<span class="va">$this</span>->charset"</span>;
        <span class="va">$options</span> = [
            <span class="cl">PDO</span>::<span class="nu">ATTR_ERRMODE</span>            => <span class="cl">PDO</span>::<span class="nu">ERRMODE_EXCEPTION</span>,
            <span class="cl">PDO</span>::<span class="nu">ATTR_DEFAULT_FETCH_MODE</span> => <span class="cl">PDO</span>::<span class="nu">FETCH_OBJ</span>,
            <span class="cl">PDO</span>::<span class="nu">ATTR_EMULATE_PREPARES</span>   => <span class="kw">false</span>,
        ];
        <span class="va">$this</span>-><span class="va">pdo</span> = <span class="kw">new</span> <span class="cl">PDO</span>(<span class="va">$dsn</span>, <span class="va">$this</span>-><span class="va">user</span>, <span class="va">$this</span>-><span class="va">pass</span>, <span class="va">$options</span>);
        <span class="kw">return</span> <span class="va">$this</span>-><span class="va">pdo</span>;
    }

    <span class="hl">    <span class="kw">public function</span> <span class="fn">safeQuery</span>(<span class="kw">string</span> <span class="va">$sql</span>, <span class="kw">array</span> <span class="va">$params</span> = []): \<span class="cl">PDOStatement</span> {</span>
<span class="hl">        <span class="va">$stmt</span> = <span class="va">$this</span>-><span class="va">pdo</span>-><span class="fn">prepare</span>(<span class="va">$sql</span>);</span>
<span class="hl">        <span class="va">$stmt</span>-><span class="fn">execute</span>(<span class="va">$params</span>);</span>
<span class="hl">        <span class="kw">return</span> <span class="va">$stmt</span>;</span>
<span class="hl">    }</span>

    <span class="kw">public function</span> <span class="fn">insert_id</span>(): <span class="kw">string</span> {
        <span class="kw">return</span> <span class="va">$this</span>-><span class="va">pdo</span>-><span class="fn">lastInsertId</span>();
    }
}</pre>
    </div>
</section>

<!-- ══════════════════════════════════════════
     EXCEL READER
════════════════════════════════════════════ -->
<section id="excel-reader">
    <div class="section-label">// 07 · EXCEL READER</div>
    <h2>Excel Reader — PHP 8 Değişiklikleri</h2>
    <p>
        Orijinal <code>excel_reader2.php</code> (PHP 4/5 dönemi) PHP 8'de aşağıdaki hatalarla çöküyordu.
        Tüm sorunlar giderildi.
    </p>

    <table class="url-table">
        <thead>
            <tr>
                <th>SORUN</th>
                <th>ESKİ (PHP 4/5)</th>
                <th>YENİ (PHP 8)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Constructor</td>
                <td><code>function OLERead()</code></td>
                <td><code>public function __construct()</code></td>
            </tr>
            <tr>
                <td>Erişim belirleyici</td>
                <td><code>var $data</code></td>
                <td><code>public string $data</code></td>
            </tr>
            <tr>
                <td>Hata bastırma</td>
                <td><code>@$sonuc = fetch()</code></td>
                <td><code>while ($row = fetch())</code></td>
            </tr>
            <tr>
                <td>Tip bildirimi</td>
                <td>Yok</td>
                <td><code>function read(string $file): bool</code></td>
            </tr>
            <tr>
                <td>Null erişim</td>
                <td><code>$Row[0]</code> (uyarı)</td>
                <td><code>$Row[0] ?? ''</code></td>
            </tr>
            <tr>
                <td>array_comb()</td>
                <td>Custom function</td>
                <td>PHP native <code>array_combine()</code></td>
            </tr>
        </tbody>
    </table>

    <div class="alert success">
        <span class="alert-icon">✅</span>
        <p>PHP 8'de <code>Deprecated</code>, <code>Warning</code> veya <code>Fatal error</code> üretmeyen temiz bir kod tabanı elde edildi.</p>
    </div>
</section>

<!-- ══════════════════════════════════════════
     GÜVENLİK
════════════════════════════════════════════ -->
<section id="guvenlik">
    <div class="section-label">// 08 · GÜVENLİK</div>
    <h2>Güvenlik Notları</h2>

    <div class="alert danger">
        <span class="alert-icon">🚨</span>
        <p>Orijinal kodda <strong>SQL Injection</strong> ve <strong>XSS</strong> açıkları mevcuttu. Bu versiyon her ikisini de tamamen kapatmaktadır.</p>
    </div>

    <h3>1 · SQL Injection Koruması</h3>
    <p>Tüm kullanıcı girdileri doğrudan sorguya eklenmek yerine <code>?</code> placeholder ile bağlanır. PDO, değerleri otomatik olarak kaçışlar.</p>

    <h3>2 · XSS Koruması</h3>
    <p>Veritabanından gelen veya Excel'den okunan her veri <code>htmlspecialchars()</code> ile HTML'e güvenli şekilde yazılır.</p>

    <h3>3 · Dosya Yükleme Güvenliği</h3>
    <p>Yalnızca izin verilen MIME tipleri kabul edilir. Dosya adı <code>basename()</code> ile temizlenir. Geçici dosya işlem sonrası <code>unlink()</code> ile silinir.</p>

    <div class="alert warn">
        <span class="alert-icon">⚠️</span>
        <p>Prodüksiyonda <code>$host</code>, <code>$user</code>, <code>$pass</code> bilgilerini kaynak kodda değil <code>.env</code> dosyasında ya da ortam değişkenlerinde saklayın.</p>
    </div>

    <h3>4 · Prodüksiyon Ortamı İçin Ek Öneriler</h3>
    <div class="code-wrap">
        <span class="code-label">php.ini — Önerilen ayarlar</span>
        <pre>display_errors  = Off
log_errors      = On
error_log       = /var/log/php_errors.log
upload_max_filesize = 10M
post_max_size       = 12M</pre>
    </div>
</section>

<!-- ══════════════════════════════════════════
     FOOTER
════════════════════════════════════════════ -->
<footer>
    <p>
        PHP 8 Excel Import · 
        <a href="https://github.com/kullanici/php8-excel-import">GitHub'da Görüntüle</a> · 
        MIT Lisansı
    </p>
    <p style="margin-top:6px; font-size:12px; color:#334155">
        Orijinal SpreadsheetReader: <a href="https://github.com/nuovo/spreadsheet-reader">nuovo/spreadsheet-reader</a> · 
        Orijinal Excel Reader: <a href="http://code.google.com/p/php-excel-reader/">php-excel-reader</a>
    </p>
</footer>

</div><!-- .wrapper -->

<script>
    // Aktif nav linkini güncelle
    const sections = document.querySelectorAll('section[id]');
    const navLinks  = document.querySelectorAll('nav a');

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                navLinks.forEach(a => a.classList.remove('active'));
                const active = document.querySelector(`nav a[href="#${entry.target.id}"]`);
                if (active) {
                    active.classList.add('active');
                    active.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
                }
            }
        });
    }, { rootMargin: '-30% 0px -60% 0px' });

    sections.forEach(s => observer.observe(s));
</script>

</body>
</html>