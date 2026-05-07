#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QListWidgetItem>

#include "characterslotwidget.h"

#define MAX_SQUAD_SIZE 5

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class SquadMember{
public:
    SquadMember();
public:
    int db_id = 0;
    QString name;
    QPixmap img;
};

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private:
    void queryAndUpdateWidget(QString sql_command);
    QString getColumnName(QString str);
    void updateSquadWidgets();
private:
    Ui::MainWindow *ui;
    QMap<QString, QStringList> map_filter;
    QList<SquadMember> squad;
    CharacterSlotWidget *widget_squad[MAX_SQUAD_SIZE];

private slots:
    void onFilterButtonToggled(bool checked);
    void onCharacterClicked(QListWidgetItem *item);
    void onSquadClicked(CharacterSlotWidget *widget_member);
};
#endif // MAINWINDOW_H
