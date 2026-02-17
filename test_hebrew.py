"""
test_hebrew.py  –  Checks Hebrew (RTL) support end-to-end.
"""
from markitdown import MarkItDown
from pathlib import Path

# ── 1. Create a Hebrew HTML test file ───────────────────────────────────────
html = """\
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head><meta charset="UTF-8"><title>דוח מכירות</title></head>
<body>
  <h1>דוח מכירות רבעוני</h1>
  <p>זהו מסמך בדיקה <strong>בעברית</strong> עבור MarkItDown.</p>

  <h2>סיכום</h2>
  <p>הביצועים ברבעון זה היו <em>מעל הציפיות</em>.</p>

  <h2>מדדים עיקריים</h2>
  <table border="1" cellpadding="4">
    <thead>
      <tr><th>אזור</th><th>הכנסות (₪)</th><th>צמיחה</th></tr>
    </thead>
    <tbody>
      <tr><td>צפון</td><td>1,200,000</td><td>+12%</td></tr>
      <tr><td>דרום</td><td>980,000</td><td>+8%</td></tr>
      <tr><td>מזרח</td><td>1,450,000</td><td>+18%</td></tr>
      <tr><td>מערב</td><td>870,000</td><td>+5%</td></tr>
    </tbody>
  </table>

  <h2>פעולות נדרשות</h2>
  <ul>
    <li>הגדלת תקציב שיווק באזור המערב</li>
    <li>השקת קו מוצרים חדש ברבעון הבא</li>
    <li>סקירת אסטרטגיית תמחור לאזור הדרום</li>
  </ul>

  <blockquote>
    <p>״ביצוע הוא הכל.״ – תזכיר פנימי</p>
  </blockquote>
</body>
</html>
"""

src = Path("input/hebrew_test.html")
src.write_text(html, encoding="utf-8")
print(f"Created: {src}")

# ── 2. Convert to Markdown ───────────────────────────────────────────────────
md_converter = MarkItDown()
result = md_converter.convert(str(src))
md_text = result.text_content

out_md = Path("output/hebrew_test.md")
out_md.write_text(md_text, encoding="utf-8")
print(f"Markdown -> {out_md}")
print("\n=== Markdown preview (char count: %d, has Hebrew: %s) ===" % (
    len(md_text),
    any('\u0590' <= c <= '\u05FF' for c in md_text)
))

# ── 3. Generate PDF ──────────────────────────────────────────────────────────
from convert import _build_pdf
out_pdf = Path("output/hebrew_test.pdf")
try:
    _build_pdf(md_text, out_pdf)
    print(f"\nPDF      -> {out_pdf}")
    print("PDF generated successfully.")
except Exception as e:
    print(f"\n[PDF ERROR] {e}")
