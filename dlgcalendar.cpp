#include <QFileDialog>
#include "dlgcalendar.h"
#include "./ui_dlgcalendar.h"
#include "include_cpp/libxl.h"
#include "date.h"
#include <chrono>

using namespace std::chrono_literals;


DlgCalendar::DlgCalendar(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::DlgCalendar)
{
    ui->setupUi(this);
    QFileDialog dlg(this, tr("Select calendar file (*.ics)"), ".", tr("Calendar Files (*.ics)"));
    if(dlg.exec())
    {
        QStringList files;
        files = dlg.selectedFiles();
        for(const auto& f : files)
        {
            import(f);
        }
        QString file = QFileDialog::getSaveFileName(this, "Select file to write", ".", tr("Excel file (*.xls)"));
        exportOnExcel(file);
    }
}

DlgCalendar::~DlgCalendar()
{
    delete ui;
}

void DlgCalendar::import(const QString& filename)
{
    enum
    {
        _idle = 0,
        _event = 1,
        _description = 2,
    }status = _idle;
    ZEvent event;

    QFile data(filename);
    if(data.open(QFile::ReadOnly | QFile::Truncate))
    {
        QTextStream imp(&data);
        QString line;
        while(imp.readLineInto(&line))
        {
            if(status == _description)
            {
                if(line.size() > 0 && line.at(0) == ' ')
                    event.description.append(line.mid(1).toStdString());
                else
                    status = _event;
            }
            QString l = line.simplified();
            switch(status)
            {
            case _idle:
                if(l.contains("BEGIN:VEVENT"))
                {
                    status = _event;
                }
                break;
            case _event:
                if(l.contains("DTSTART:"))
                {
                    std::tm tm = {};
                    std::stringstream ss(l.mid(8, 13).toStdString());
                    ss >> std::get_time(&tm, "%Y%m%dT%H%M%S");
                    event.start = std::chrono::system_clock::from_time_t(std::mktime(&tm));
                }
                else if(l.contains("DTEND:"))
                {
                    std::tm tm = {};
                    std::stringstream ss(l.mid(6, 13).toStdString());
                    ss >> std::get_time(&tm, "%Y%m%dT%H%M%S");
                    event.end = std::chrono::system_clock::from_time_t(std::mktime(&tm));
                    event.duration = event.end - event.start;
                }
                else if(l.contains("DESCRIPTION:"))
                {
                    event.description = l.mid(12).toStdString();
                    status = _description;
                }
                else if(l.contains("SUMMARY:"))
                {
                    event.title = l.mid(8).toStdString();
                    event.code = std::atoi(event.title.data() + 1);
                    event.travel = (event.code > 0 && event.title.length() > 0 && (event.title[0] == 'V' || event.title[0] == 'v'));
                }
                else if(l.contains("END:VEVENT") && (event.title[0] == 'C' || event.title[0] == 'V') && event.code > 0)
                {
                    m_events.insert(m_events.begin(), event);
                    status = _idle;
                }
                break;
            default:
                break;
            }
        }
    }

}
void DlgCalendar::exportOnExcel(const QString& filename)
{
    using namespace libxl;
    wchar_t name[256];
    int len;

    Book* book = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files

    Font* font = book->addFont();
    font->setColor(COLOR_RED);
    font->setBold(true);
    Format* dateFormat = book->addFormat();
    dateFormat->setNumFormat(NUMFORMAT_DATE);
    Format* timeFormat = book->addFormat();
    timeFormat->setNumFormat(NUMFORMAT_CUSTOM_HMMSS);

    double difftime = 3.; // difference between local time and utc

    Sheet* sheet = book->addSheet(L"Timesheet");
    sheet->writeStr(1, 0, L"Foglio ore Paolo Liprandi !");
    sheet->writeStr(2, 0, L"ID");
    sheet->writeStr(2, 1, L"Data");
    sheet->writeStr(2, 2, L"Commessa");
    sheet->writeStr(2, 3, L"Inizio");
    sheet->writeStr(2, 4, L"Fine");
    sheet->writeStr(2, 5, L"Pausa");
    sheet->writeStr(2, 6, L"Durata");
    sheet->writeStr(2, 7, L"Titolo");
    sheet->writeStr(2, 8, L"Descrizione");
    int row = 3;
    for(const auto& ev : m_events)
    {
        auto dp = std::chrono::floor<date::days>(ev.start);
        auto sod = ev.start - dp;
        auto eod = ev.end - dp;
        date::year_month_day ymd{dp};
        date::hh_mm_ss time{std::chrono::floor<std::chrono::milliseconds>(ev.start - dp)};
        date::weekday wd{dp};
        sheet->writeNum(row, 0, row - 3);
        // 25569 is the number of days between 1 January 1900 and 1 January 1970
        sheet->writeNum(row, 1, dp.time_since_epoch().count() + 25569, dateFormat);
        if(!ev.travel)
            sheet->writeNum(row, 2, ev.code);
        else
            sheet->writeStr(row, 2, L"VIAGGIO");
        sheet->writeNum(row, 3, sod / 1.0s / (24. * 3600.) + difftime / 24., timeFormat);
        sheet->writeNum(row, 4, eod / 1.0s / (24. * 3600.) + difftime / 24., timeFormat);
        double seconds = ev.duration / 1.0s;
        sheet->writeNum(row, 5, (seconds > 6. * 3600) ? 1./24.: 0, timeFormat);
        if(seconds > 6. * 3600)
            seconds -= 3600;
        sheet->writeNum(row, 6, seconds / (24. * 3600.), timeFormat);
        QString str = QString::fromStdString(ev.title);
        str.truncate(250);
        len = str.toWCharArray(name);
        name[len] = 0;
        sheet->writeStr(row,7, name);
        str = QString::fromStdString(ev.description);
        str.truncate(250);
        len = str.toWCharArray(name);
        name[len] = 0;
        sheet->writeStr(row, 8, name);
        row++;
    }


    len = filename.toWCharArray(name);
    name[len] = 0;
    book->save(name);

    book->release();
 }
