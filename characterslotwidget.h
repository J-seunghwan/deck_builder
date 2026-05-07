#ifndef CHARACTERSLOTWIDGET_H
#define CHARACTERSLOTWIDGET_H

#include <QFrame>
#include <QMouseEvent>

namespace Ui {
class CharacterSlotWidget;
}

class CharacterSlotWidget : public QFrame
{
    Q_OBJECT

public:
    explicit CharacterSlotWidget(QWidget *parent = nullptr);
    ~CharacterSlotWidget();

    int getCharacterId();
    void setCharacterId(int num);

    bool getCharacterClicked();
    void setCharacterClicked(bool check);

    void setCharacterProfileInfo(QString name, QPixmap img);
    void setCharacterProfileInfo(QString name, QString path);
    QString getCharacterName();
    QPixmap getCharacterImg();
    void setupInitState();
    QString printInfo();

private:
    Ui::CharacterSlotWidget *ui;//CharacterSlot.ui 용
    int db_id = 0;//db 쿼리 용.
    bool checked = false;

signals:
    void clicked(CharacterSlotWidget* widget);

protected:
    void mousePressEvent(QMouseEvent *event) override;
};

#endif // CHARACTERSLOTWIDGET_H
