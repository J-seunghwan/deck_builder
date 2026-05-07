#include "characterslotwidget.h"
#include "ui_character_slot_widget.h"

CharacterSlotWidget::CharacterSlotWidget(QWidget *parent)
    : QFrame(parent), ui(new Ui::CharacterSlotWidget)
{
    this->ui->setupUi(this);
    this->setupInitState();
}

CharacterSlotWidget::~CharacterSlotWidget(){
    delete this->ui;
}

int CharacterSlotWidget::getCharacterId(){
    return this->db_id;
}

void CharacterSlotWidget::setCharacterId(int num){
    this->db_id = num;
}

bool CharacterSlotWidget::getCharacterClicked(){
    return this->checked;
}

void CharacterSlotWidget::setCharacterClicked(bool check){
    this->checked = check;
}

void CharacterSlotWidget::setCharacterProfileInfo(QString name, QPixmap img){
    this->ui->label_char_name->setText(name);
    this->ui->label_profile_img->setPixmap(img);
}

void CharacterSlotWidget::setCharacterProfileInfo(QString name, QString path){
    this->ui->label_char_name->setText(name);

    QPixmap pixmap(path);
    if(!pixmap.isNull()){
        this->ui->label_profile_img->setPixmap(pixmap);
    }
}

void CharacterSlotWidget::setupInitState(){
    this->db_id = 0;
    this->ui->label_char_name->setText(QString(""));
    this->ui->label_profile_img->setPixmap(QPixmap(QString("resource/empty_slot_img.png")));
}

QString CharacterSlotWidget::getCharacterName(){
    return this->ui->label_char_name->text();
}

QPixmap CharacterSlotWidget::getCharacterImg(){
    return this->ui->label_profile_img->pixmap();
}

QString CharacterSlotWidget::printInfo(){
    return QString("name %1 db_id %2 ui_clicked %3").arg(this->getCharacterName()).arg(this->getCharacterId()).arg(this->getCharacterClicked());
}

void CharacterSlotWidget::mousePressEvent(QMouseEvent *event){
    if(event->button() == Qt::LeftButton){
        emit clicked(this);
    }
    QFrame::mousePressEvent(event);
}