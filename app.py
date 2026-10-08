import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import cleaner


# ============================================================
# إعدادات التصميم
# ============================================================

BG_COLOR = "#F4F6F8"
CARD_COLOR = "#FFFFFF"
PRIMARY_COLOR = "#1F4E78"
PRIMARY_HOVER = "#163A5C"
SUCCESS_COLOR = "#198754"
TEXT_COLOR = "#263238"
SECONDARY_TEXT = "#607D8B"
BORDER_COLOR = "#D9E1E8"


# ============================================================
# وظائف التطبيق
# ============================================================

def update_progress(value, message):
    """تحديث شريط التقدم والنص بأمان من Thread الخلفية."""
    window.after(0, lambda: apply_progress(value, message))


def apply_progress(value, message):
    """تحديث واجهة المستخدم بالنسبة الحالية."""
    progress_bar["value"] = value
    progress_percent_label.config(text=f"{int(value)}%")
    status_label.config(text=message, fg=PRIMARY_COLOR)


def run_cleaner():
    """بدء عملية تنظيف البيانات في Thread منفصل."""

    run_button.config(state="disabled")
    open_excel_button.config(state="disabled")

    report_label.config(text="جاري تنفيذ عملية التنظيف...")

    progress_bar["value"] = 0
    progress_percent_label.config(text="0%")

    status_label.config(
        text="جاري بدء عملية التنظيف...",
        fg=PRIMARY_COLOR
    )

    progress_frame.pack(
        padx=40,
        pady=(0, 20),
        fill="x"
    )

    cleaning_thread = threading.Thread(
        target=perform_cleaning,
        daemon=True
    )
    cleaning_thread.start()


def perform_cleaning():
    """تشغيل cleaner.py في الخلفية."""

    try:
        cleaning_stats = cleaner.main(
            progress_callback=update_progress
        )

        window.after(
            0,
            cleaning_finished,
            cleaning_stats,
            None
        )

    except Exception as error:
        window.after(
            0,
            cleaning_finished,
            None,
            error
        )


def cleaning_finished(cleaning_stats, error):
    """معالجة نتيجة عملية التنظيف وتحديث الواجهة."""

    if error is not None:
        # إعادة شريط التقدم إلى الصفر عند حدوث خطأ
        progress_bar["value"] = 0
        progress_percent_label.config(text="0%")

        status_label.config(
            text="فشلت عملية التنظيف.",
            fg="#DC3545"
        )

        report_label.config(
            text="فشلت عملية التنظيف.\n"
                 "راجع ملف cleaner.log لمعرفة التفاصيل."
        )

        messagebox.showerror(
            "فشل تنظيف البيانات",
            "تعذر إكمال عملية تنظيف البيانات.\n\n"
            "راجع ملف cleaner.log لمعرفة التفاصيل."
        )

    elif cleaning_stats:
        progress_bar["value"] = 100
        progress_percent_label.config(text="100%")

        status_label.config(
            text="اكتملت عملية التنظيف بنجاح.",
            fg=SUCCESS_COLOR
        )

        report_text = (
            "تم تنظيف البيانات بنجاح.\n\n"
            f"الصفوف قبل التنظيف: "
            f"{cleaning_stats['rows_before']}\n"
            f"الصفوف المكررة المحذوفة: "
            f"{cleaning_stats['duplicates_removed']}\n"
            f"الصفوف ذات القيم الفارغة المحذوفة: "
            f"{cleaning_stats['missing_rows']}\n"
            f"الإيميلات غير الصحيحة المحذوفة: "
            f"{cleaning_stats['invalid_emails_removed']}\n"
            f"الصفوف بعد التنظيف: "
            f"{cleaning_stats['rows_after']}"
        )

        report_label.config(text=report_text)

        open_excel_button.config(state="normal")

    else:
        # التأكد من سبب الخطأ من ملف الـ log لعرض رسالة مناسبة
        excel_file_error = False

        try:
            log_path = cleaner.LOG_FILE

            if os.path.exists(log_path):
                with open(
                    log_path,
                    "r",
                    encoding="utf-8"
                ) as log_file:
                    log_content = log_file.read()

                excel_file_error = (
                    "فشل حفظ Excel لأن الملف مفتوح أو محمي"
                    in log_content
                )

        except Exception:
            excel_file_error = False

        # إعادة شريط التقدم إلى الصفر
        progress_bar["value"] = 0
        progress_percent_label.config(text="0%")

        status_label.config(
            text="فشلت عملية التنظيف.",
            fg="#DC3545"
        )

        if excel_file_error:
            report_label.config(
                text="تعذر حفظ ملف Excel.\n\n"
                     "ملف customers_filtered.xlsx مفتوح في Excel.\n"
                     "أغلق الملف ثم اضغط تشغيل تنظيف البيانات مرة أخرى."
            )

            messagebox.showerror(
                "ملف Excel مفتوح",
                "تعذر حفظ ملف Excel لأن:\n\n"
                "customers_filtered.xlsx مفتوح في Excel.\n\n"
                "أغلق الملف ثم اضغط تشغيل تنظيف البيانات مرة أخرى."
            )

        else:
            report_label.config(
                text="فشلت عملية التنظيف.\n"
                     "راجع ملف cleaner.log لمعرفة التفاصيل."
            )

            messagebox.showerror(
                "خطأ",
                "فشل تنظيف البيانات.\n"
                "راجع ملف cleaner.log لمعرفة التفاصيل."
            )

    # إعادة تفعيل زر التشغيل بعد انتهاء العملية
    run_button.config(state="normal")


def open_excel_file():
    """فتح ملف Excel الناتج."""

    file_name = cleaner.FILE_NAME

    if os.path.exists(file_name):
        try:
            os.startfile(file_name)
        except Exception as error:
            messagebox.showerror(
                "خطأ",
                f"تعذر فتح ملف Excel:\n\n{error}"
            )
    else:
        messagebox.showerror(
            "خطأ",
            "ملف Excel غير موجود."
        )


# ============================================================
# إنشاء النافذة
# ============================================================

window = tk.Tk()
window.title("Excel Data Cleaner")
window.geometry("650x650")
window.resizable(False, False)
window.configure(bg=BG_COLOR)


# ============================================================
# العنوان
# ============================================================

title_label = tk.Label(
    window,
    text="Excel Data Cleaner",
    font=("Arial", 24, "bold"),
    bg=BG_COLOR,
    fg=PRIMARY_COLOR
)
title_label.pack(pady=(35, 5))


subtitle_label = tk.Label(
    window,
    text="تنظيف وتنظيم بيانات Excel تلقائيًا",
    font=("Arial", 11),
    bg=BG_COLOR,
    fg=SECONDARY_TEXT
)
subtitle_label.pack(pady=(0, 25))


# ============================================================
# كارت التحكم
# ============================================================

control_card = tk.Frame(
    window,
    bg=CARD_COLOR,
    highlightbackground=BORDER_COLOR,
    highlightthickness=1
)
control_card.pack(padx=40, fill="x")


run_button = tk.Button(
    control_card,
    text="تشغيل تنظيف البيانات",
    font=("Arial", 13, "bold"),
    bg=PRIMARY_COLOR,
    fg="white",
    activebackground=PRIMARY_HOVER,
    activeforeground="white",
    relief="flat",
    bd=0,
    cursor="hand2",
    command=run_cleaner
)
run_button.pack(padx=25, pady=(25, 10), fill="x")


open_excel_button = tk.Button(
    control_card,
    text="فتح ملف Excel",
    font=("Arial", 12),
    bg="#E9EEF3",
    fg=TEXT_COLOR,
    activebackground="#DCE5EC",
    activeforeground=TEXT_COLOR,
    relief="flat",
    bd=0,
    cursor="hand2",
    command=open_excel_file,
    state="disabled"
)
open_excel_button.pack(padx=25, pady=(0, 25), fill="x")


# ============================================================
# الحالة
# ============================================================

status_label = tk.Label(
    window,
    text="جاهز لتنظيف البيانات",
    font=("Arial", 11, "bold"),
    bg=BG_COLOR,
    fg=SECONDARY_TEXT
)
status_label.pack(pady=(25, 10))


# ============================================================
# شريط التقدم + النسبة المئوية
# ============================================================

progress_frame = tk.Frame(
    window,
    bg=BG_COLOR
)

progress_percent_label = tk.Label(
    progress_frame,
    text="0%",
    font=("Arial", 11, "bold"),
    bg=BG_COLOR,
    fg=PRIMARY_COLOR
)
progress_percent_label.pack(pady=(0, 6))


progress_bar = ttk.Progressbar(
    progress_frame,
    orient="horizontal",
    mode="determinate",
    maximum=100,
    value=0
)
progress_bar.pack(fill="x")

progress_frame.pack_forget()


# ============================================================
# كارت التقرير
# ============================================================

report_card = tk.Frame(
    window,
    bg=CARD_COLOR,
    highlightbackground=BORDER_COLOR,
    highlightthickness=1
)
report_card.pack(
    padx=40,
    pady=10,
    fill="both",
    expand=True
)


report_title = tk.Label(
    report_card,
    text="تقرير التنظيف",
    font=("Arial", 14, "bold"),
    bg=CARD_COLOR,
    fg=PRIMARY_COLOR
)
report_title.pack(pady=(20, 10))


report_label = tk.Label(
    report_card,
    text="سيظهر تقرير التنظيف هنا",
    font=("Arial", 11),
    bg=CARD_COLOR,
    fg=TEXT_COLOR,
    justify="left"
)
report_label.pack(padx=25, pady=10)


# ============================================================
# تشغيل التطبيق
# ============================================================

window.mainloop()
