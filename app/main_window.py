import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from app import constants
from service.average_workbook import write_average_workbook
from service.distribution_csv import write_distribution_csv
from service.mail_merge_csv import write_mail_merge_csv
from service.revenue_reader import parse_target_month, read_doctor_revenues
from service.trend_workbook import backup_workbook, update_trend_workbook
from utils.config_manager import ConfigManager

logger = logging.getLogger(__name__)


class MainWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.config_manager = ConfigManager()
        self.revenue_path = tk.StringVar()

        root.title(constants.WINDOW_TITLE)
        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=constants.LABEL_REVENUE_FILE).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.revenue_path, width=60).grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(frame, text=constants.BUTTON_BROWSE, command=self._select_file).grid(row=0, column=2)
        ttk.Button(frame, text=constants.BUTTON_RUN, command=self._run).grid(row=1, column=0, columnspan=3, pady=10)
        self.result_text = scrolledtext.ScrolledText(frame, height=15, state=tk.DISABLED)
        self.result_text.grid(row=2, column=0, columnspan=3, sticky=tk.NSEW)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(2, weight=1)

    def _select_file(self) -> None:
        selected = filedialog.askopenfilename(
            title=constants.DIALOG_TITLE_SELECT_FILE,
            initialdir=self.config_manager.get_path("trend_workbook").parent,
            filetypes=[(constants.FILE_TYPE_EXCEL, "*.xlsx")],
        )
        if selected:
            self.revenue_path.set(selected)

    def _run(self) -> None:
        if not self.revenue_path.get():
            messagebox.showerror(constants.TITLE_ERROR, constants.MSG_NO_FILE_SELECTED)
            return
        try:
            month, inserted_rows = self._aggregate(Path(self.revenue_path.get()))
        except PermissionError as e:
            self._show_error(constants.MSG_FILE_IN_USE, e)
        except FileNotFoundError as e:
            self._show_error(constants.MSG_FILE_NOT_FOUND, e)
        except Exception as e:
            self._show_error(constants.MSG_UNEXPECTED_ERROR, e)
        else:
            message = constants.MSG_DONE.format(month=month)
            if inserted_rows:
                message += constants.MSG_INSERTED_ROWS.format(count=len(inserted_rows), rows="\n".join(inserted_rows))
            messagebox.showinfo(constants.TITLE_DONE, message)

    def _aggregate(self, revenue_path: Path) -> tuple[int, list[str]]:
        """集計を実行し、対象月と変化表に追加した行の通知文を返す"""
        trend_path = self.config_manager.get_path("trend_workbook")
        csv_path = self.config_manager.get_path("distribution_csv")
        average_path = self.config_manager.get_path("average_workbook")
        mail_merge_path = self.config_manager.get_path("mail_merge_csv")

        month = parse_target_month(revenue_path)
        revenues = read_doctor_revenues(revenue_path)
        backup_dir = self.config_manager.get_path("backup_dir")
        backup_generations = self.config_manager.config.getint("Backup", "generations")
        self._report(constants.LOG_BACKUP.format(path=backup_workbook(trend_path, backup_dir, backup_generations)))
        inserted = update_trend_workbook(trend_path, month, revenues)
        self._report(constants.LOG_WORKBOOK_UPDATED.format(month=month, count=len(revenues), path=trend_path))
        inserted_rows: list[str] = []
        for row, doctor in inserted:
            self._report(constants.LOG_DOCTOR_INSERTED.format(
                row=row, department=doctor.department, doctor_id=doctor.doctor_id, name=doctor.name))
            inserted_rows.append(constants.MSG_INSERTED_ROW.format(
                row=row, department=doctor.department, doctor_id=doctor.doctor_id, name=doctor.name))

        recipients_path = self.config_manager.get_path("recipients_csv")
        missing_ids = write_distribution_csv(trend_path, recipients_path, csv_path)
        self._report(constants.LOG_CSV_WRITTEN.format(path=csv_path))
        for doctor_id in missing_ids:
            self._report(constants.LOG_MISSING_RECIPIENT.format(doctor_id=doctor_id))
        write_average_workbook(trend_path, recipients_path, average_path)
        self._report(constants.LOG_AVERAGE_WRITTEN.format(path=average_path))
        write_mail_merge_csv(trend_path, recipients_path, mail_merge_path)
        self._report(constants.LOG_MAIL_MERGE_WRITTEN.format(path=mail_merge_path))
        return month, inserted_rows

    def _report(self, message: str) -> None:
        logger.info(message)
        self.result_text.configure(state=tk.NORMAL)
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.configure(state=tk.DISABLED)
        self.result_text.see(tk.END)

    def _show_error(self, template: str, error: Exception) -> None:
        logger.exception(error)
        messagebox.showerror(constants.TITLE_ERROR, template.format(detail=error))
