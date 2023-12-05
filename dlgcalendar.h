#ifndef DLGCALENDAR_H
#define DLGCALENDAR_H

#include <QDialog>
#include <iostream>
#include <chrono>

QT_BEGIN_NAMESPACE
namespace Ui { class DlgCalendar; }
QT_END_NAMESPACE
class ZEvent
{
public:
    ZEvent():
        duration(0)
      ,  code(0)
      , travel(false)
      , title("")
      , description("")
    {}
public:
    std::chrono::time_point<std::chrono::system_clock> start;   // start date and time
    std::chrono::time_point<std::chrono::system_clock> end;     // end date and time
    std::chrono::duration<double> duration;                     // duration removing lanch break time
    int code;                                                   // order number
    bool travel;                                                // if true is travel time
    std::string title;                                          // title
    std::string description;                                    // description
};

class DlgCalendar : public QDialog
{
    Q_OBJECT

public:
    DlgCalendar(QWidget *parent = nullptr);
    ~DlgCalendar();

private:
    void import(const QString& filename);
    void exportOnExcel(const QString& filename);
private:
    Ui::DlgCalendar *ui;
    std::list<ZEvent> m_events;
};
#endif // DLGCALENDAR_H
