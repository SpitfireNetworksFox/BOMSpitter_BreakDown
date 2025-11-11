# quote_html_template.py

quoteTemplateHtml = r"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>{{ bom["bom_name"] }}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@200..800&display=swap" rel="stylesheet">
<style>
/* Tighter page margins for PDF; page number footer stays the same */
@page {
  size: Letter;
  margin: 0.2in; /* was 0.3in */
  @bottom-right {
    content: "{{ quote_number }} " counter(page) "/" counter(pages);
    font: 700 12px Manrope, sans-serif;
  }
}

*{ box-sizing:border-box }
body{ margin:0; padding:20px; font-family:Manrope,sans-serif; background:#f8f9fa }

/* Main container */
.page-container{
  max-width:8.5in;
  margin:0 auto;
  background:white;
  padding:25px;
  box-shadow:0 0 20px rgba(0,0,0,0.1);
  min-height:11in;
}

/* Header */
.header-container{
  position:relative;
  margin-bottom:20px;
  padding:0 220px 15px 0; /* reserve space on right for quote number */
  border-bottom:2px solid #061528;
}
.header-container h4,.header-container p{ margin:4px 0; line-height:1.3 }
.header-container h4{ font-size:16px; font-weight:600; color:#333 }
.header-container p{ font-size:12px; color:#666 }

.logo {
  position: absolute;
  top: 40px;              /* 1 inch down */
  left: 50%;
  transform: translateX(-50%);
  width: auto;
  height: auto;
  max-width: 200px;
}

.q_num{ position:absolute; top:0; right:0; font-size:22px; font-weight:700; color:#333 }
.term{  position:absolute; top:25px; right:0; font-size:16px; color:#666 }

/* Terms table */
.terms-table{
  width:100%;
  border-collapse:collapse;
  margin:15px 0;
  border:2px solid #061528;
  background:white;
}
.terms-table thead th{
  background:#061528; color:white; font-size:11px; font-weight:600; text-align:center;
  padding:12px 8px; border-right:1px solid #000000;
}
.terms-table tbody td{
  font-size:11px; text-align:center; padding:12px 8px; border-right:1px solid #ddd; background:#f9f9f9;
}

/* Items table */
.items-section{ 
  margin:20px 0; 
  page-break-inside: auto; /* Allow breaking inside items section */
}
.items-table{
  width:100%; border-collapse:collapse; border:2px solid #061528; margin-bottom:15px;
}
.items-table thead th{
  background:#061528; color:white; font-size:11px; font-weight:600; text-align:center;
  padding:12px 8px; border-right:1px solid #000000;
}
.items-table tbody tr{ border-bottom:1px solid #ddd; }
.items-table tbody tr:nth-child(even){ background:#f9f9f9; }
.items-table tbody tr:hover{ background:#f0f8ff; }
.items-table td{
  font-size:10px; padding:10px 8px; border-right:1px solid #ddd; vertical-align:top;
}

/* Column widths + alignment */
.col-item     { width:40px;  text-align:center; }
.col-cat      { width:110px; text-align:center; }
.col-sku      { width:120px; text-align:center; }
.col-qty      { width:40px;  text-align:center; }
.col-desc { 
  width:300px; 
  text-align:left; 
  word-wrap:break-word; 
  line-height:1.3; 
  white-space: pre-wrap;   /* ← add this */
}
.col-price    { width:80px;  text-align:right; }
.col-extended { width:80px;  text-align:right; }

/* Total section - separate from table */
.total-section{
  margin-top:15px;
  border:2px solid #061528;
  background:white;
  page-break-inside: avoid; /* Prevent breaking inside the total section */
  break-inside: avoid; /* For newer browsers */
  min-height: 60px; /* Ensure minimum height */
}
.total-row{ 
  background:#061528 !important; 
  color:white; 
  font-weight:700; 
  display:flex;
  align-items:center;
  min-height: 60px; /* Match the section height */
}
.total-label{
  flex:1;
  padding:20px 8px; /* Increased padding from 15px to 20px */
  font-size:12px;
  text-align:left; /* Changed from right to left */
}
.total-amount{
  min-width: 110px;   /* increased from 80px */
  padding:20px 12px;  /* a bit more padding */
  font-size:14px;     /* optional: make it larger */
  text-align:right;   /* align right looks cleaner for numbers */
  border-left:1px solid #000000;
  word-break: keep-all; /* prevent breaking mid-number */
}

/* Notes */
.notes-section{ margin:20px 0; border:2px solid #061528; background:white; }
.notes-content{ padding:20px; font-size:11px; line-height:1.4; color:#333; }
.notes-content ul{ margin:0; padding-left:20px; }
.notes-content li{ margin-bottom:8px; }

/* Signature */
/* >>> CHANGE: make signature a new-page block and unbreakable <<< */
.signature-section{
  margin-top:25px;
  border:2px solid #061528;
  background:white;

  /* never split the signature */
  break-inside: avoid;
  page-break-inside: avoid;

  /* always start the signature on a fresh page */
  break-before: page;          /* modern */
  page-break-before: always;   /* legacy */
}
.sig-header{
  background:#EDEFF7; padding:12px 15px; border-bottom:1px solid #061528; font-size:11px;
  display:flex; justify-content:space-between; align-items:flex-start;
}
.sig-contact{ flex:1; }
.sig-instructions{ padding:15px; font-size:12px; border-bottom:1px solid #ddd; background:#f9f9f9; }
.sig-fields{ padding:15px; }
.sig-line{ border-bottom:2px solid #061528; margin:25px 0 8px 0; padding-bottom:5px; }
.sig-row{ display:flex; gap:40px; margin-top:25px; }
.sig-name{ flex:2; }
.sig-date{ flex:1; }

/* Print: tighter spacing and ensure Q# stays pinned */
@media print {
  body{ background:white; padding:0; }
  .page-container{
    box-shadow:none;
    padding:12px 14px 16px 14px; /* was 25px */
    min-height:auto;             /* no forced height */
  }
  .header-container{
    padding:0 220px 8px 0; /* keep space for Q# on the right; tighter bottom */
    margin-bottom:12px;
  }
  .header-container h4{ margin:0 0 4px 0; }
  .header-container p{  margin:2px 0; }
  .q_num{ position:absolute !important; top:0; right:0; }
  .term { position:absolute !important; top:25px; right:0; }
  .items-table tbody tr:hover{ background:transparent; }

  /* keep totals as a single block */
  .total-section{
    page-break-inside: avoid !important;
    break-inside: avoid !important;
    margin-top: 10px;
    min-height: 60px !important;
  }

  .total-row{ min-height: 60px !important; }
  .total-label, .total-amount{
    padding: 20px 8px !important;
    text-align: left !important;
  }

  .items-table { margin-bottom: 10px; }
}

/* Screen-only responsive tweaks */
@media screen and (max-width: 700px) {
  .page-container{ margin:0 20px; padding:20px; }
  .header-container{ flex-direction:column; padding-right:0; }
  .q_num, .term{ position:static; text-align:left; margin-top:10px; }
}
</style>
</head>
<body>
<div class="page-container">
  <!-- Header Section -->
  <div class="header-container">
    <img class="logo" src="data:image/png;base64,{{ img_str }}" alt="Logo">
    <h4>Quotation: {{ bom["bom_name"] }}</h4>
    <h4>Spitfire Networks Inc</h4>
    <p>Attention: {{ bom["contact_name"] }}</p>
    <p>{{ bom["comp_name"] }}</p>
    <p>{{ bom["comp_address"] }}</p>
    <p>{{ bom["comp_city"] }}{% if bom["comp_city"] and (bom["comp_state"] or bom["comp_zip"]) %}, {% endif %}{{ bom["comp_state"] }} {{ bom["comp_zip"] }}</p>
    <p>{{ bom["comp_phone"] }}</p>
    <div class="q_num">{{ quote_number }}</div>
    <div class="term">Term: {{ bom["term"] }}</div>
  </div>

  <!-- Terms Section -->
  <table class="terms-table">
    <thead>
      <tr>
        <th>DATE:</th><th>TERMS:</th><th>VALID UNTIL:</th><th>INCOTERMS:</th><th>FUNDS:</th><th>DUTY:</th><th>ALL TAXES:</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>{{ date }}</td>
        <td>{{ payment_terms }}</td>
        <td>{{ expiry }}</td>
        <td>{{ bom["incoterms"] }}</td>
        <td>{{ currency }}</td>
        <td>{{ bom["duty"] }}</td>
        <td>{{ bom["taxes"] }}</td>
      </tr>
    </tbody>
  </table>

  <!-- Items Section -->
  <div class="items-section">
    <table class="items-table">
      <thead>
        <tr>
          <th class="col-item">Item #</th>
          <th class="col-cat">Category</th>
          <th class="col-sku">SKU</th>
          <th class="col-qty">Qty</th>
          <th class="col-desc">Description</th>

          {% if round_values == "True" %}
            <th class="col-price">Price</th>
          {% else %}
            {% if show_list_price == "True" %}
              <th class="col-price">List Price</th>
            {% endif %}
            <th class="col-price">Unit Price</th>
          {% endif %}
          <th class="col-extended">Extended</th>
        </tr>
      </thead>
      <tbody>
      {% for row in items %}
        <tr>
          <td class="col-item">{{ loop.index }}</td>
          <td class="col-cat">{{ row["category"] }}</td>
          <td class="col-sku">{{ row["pt_sku"] }}</td>
          <td class="col-qty">{{ row["qty"] }}</td>
          <td class="col-desc">{{ row["description"] }}</td>
        {% if round_values == "True" %}
          <td class="col-price">${{"{:,.0f}".format((row["subtotal"]|float) / (row["qty"]|float if row["qty"] else 1))}}</td>
          <td class="col-extended">${{"{:,.0f}".format(row["subtotal"]|float)}}</td>
        {% else %}
          {% if show_list_price == "True" %}
            <td class="col-price">${{"{:,.2f}".format((row["list_price"]|float) if row.get("list_price") else ((row["subtotal"]|float) / (row["qty"]|float if row["qty"] else 1)))}}</td>
          {% endif %}
          <td class="col-price">${{"{:,.2f}".format((row["subtotal"]|float) / (row["qty"]|float if row["qty"] else 1))}}</td>
          <td class="col-extended">${{"{:,.2f}".format(row["subtotal"]|float)}}</td>
        {% endif %}
        </tr>
        {% if row["notes"] %}
        <tr>
          <td colspan="4"></td>
          <td class="col-desc" style="font-style:italic;color:#666;padding-left:20px;">{{ row["notes"] }}</td>
          <td colspan="{% if round_values == 'True' %}2{% else %}{% if show_list_price == 'True' %}3{% else %}2{% endif %}{% endif %}"></td>
        </tr>
        {% endif %}
      {% endfor %}
      </tbody>
    </table>
    
    <!-- Total section - separate from table to prevent repetition -->
    <div class="total-section">
      <div class="total-row">
        <div class="total-label">Total</div>
        <div class="total-amount">
          {% if round_values == "True" %}{{"${:,.0f}".format(bom["total"]|float)}}
          {% else %}{{"${:,.2f}".format(bom["total"]|float)}}{% endif %}
        </div>
      </div>
    </div>
  </div>

  <!-- Notes Section -->
  {% if quote_notes or add_notes %}
  <div class="notes-section">
    <div class="notes-content">
      {% if quote_notes %}{{ quote_notes|safe }}
      {% elif add_notes %}
      <strong>STANDARD QUOTATION NOTES:</strong>
      <ul>
        {% for note in add_notes %}{% if note %}<li>{{note}}</li>{% endif %}{% endfor %}
      </ul>
      {% endif %}
    </div>
  </div>
  {% endif %}

  <!-- Signature Section (now always starts on a new page and never splits) -->
  <div class="signature-section">
    <div class="sig-header">
      <div class="sig-contact">
        <strong>Please Send orders to:</strong><br>
        Spitfire Networks Inc<br>
        {{ sales_desk_email }}
      </div>
      <div></div>
      <div class="sig-contact" style="text-align:right;">
        <strong>Spitfire Contact:</strong> {{ owner["name"] }}<br>
        {{ owner["email"] }}<br>
        {{ owner_phone }}
      </div>
    </div>

    <div class="sig-instructions">
      <strong>Sign Below and include this signed quotation with purchase order.</strong>
    </div>

    <div class="sig-fields">
      <div class="sig-line"></div>
      <p style="margin:5px 0;font-size:11px;color:#666;">Authorized Signature</p>

      <div class="sig-row">
        <div class="sig-name">
          <div class="sig-line"></div>
          <p style="margin:5px 0;font-size:11px;color:#666;">Print Name and Title:</p>
        </div>
        <div class="sig-date">
          <div class="sig-line"></div>
          <p style="margin:5px 0;font-size:11px;color:#666;">Date:</p>
        </div>
      </div>
    </div>
  </div>
</div>
</body></html>
"""
