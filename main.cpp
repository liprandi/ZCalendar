#include "dlgcalendar.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    DlgCalendar w;
    w.show();
    return a.exec();
}
