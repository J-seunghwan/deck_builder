#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>

#include "mainwindow.h"
#include "ui_mainwindow.h"


SquadMember::SquadMember(){

}

MainWindow::MainWindow(QWidget *parent): QMainWindow(parent), ui(new Ui::MainWindow)
{
    this->ui->setupUi(this);
    this->ui->listWidget_char->setGridSize(QSize(128, 256+50));

    this->widget_squad[0] =this->ui->widget_squad1;
    this->widget_squad[1] =this->ui->widget_squad2;
    this->widget_squad[2] =this->ui->widget_squad3;
    this->widget_squad[3] =this->ui->widget_squad4;
    this->widget_squad[4] =this->ui->widget_squad5;

    // DB 연결
    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("nikke_info.db");

    if(!db.open()){
        qDebug() << "DB 연결 실패:" << db.lastError().text();
        //todo - 팝업창으로 알리기
        return;
    }

    // DB에 등록된 캐릭터 목록 생성
    QString sql_command = "SELECT id,Name,Path_img_profile FROM nikke_info_table";
    this->queryAndUpdateWidget(sql_command);

    // 스킬들 필터용 버튼 동적 생성
    sql_command = "SELECT keyword_name, category FROM skill_keywords_table";
    QSqlQuery query;
    if(query.exec(sql_command)){
        while(query.next()){
            QString keyword_name = query.value(0).toString();
            QString category = query.value(1).toString();

            // 표시용 텍스트 ( [원문: ...] 제거 및 증가/감소 기호 변경)
            QString display_text = keyword_name;
            if(display_text.contains(" [원문:")) {
                display_text = display_text.split(" [원문:").first();
            }
            display_text.replace("증가", "▲");
            display_text.replace("감소", "▼");

            //버튼 생성 및 설정
            QPushButton *btn = new QPushButton(display_text, this);
            btn->setCheckable(true);
            btn->setProperty("original_keyword", keyword_name); // 원본 키워드 저장

            QGroupBox *groupBox = nullptr;
            if(category == "공격") groupBox = ui->groupBox_attack;
            else if(category == "보호") groupBox = ui->groupBox_protection;
            else if(category == "특수 능력") groupBox = ui->groupBox_mechanics;
            else if(category == "제어") groupBox = ui->groupBox_control;
            else if(category == "유틸리티") groupBox = ui->groupBox_utility;
            else if(category == "발동 조건") groupBox = ui->groupBox_trigger;

            if(groupBox){
                if(!groupBox->layout()){
                    QVBoxLayout *layout = new QVBoxLayout(groupBox);
                    layout->setSpacing(3);
                    layout->setContentsMargins(5, 5, 5, 5);
                }
                groupBox->layout()->addWidget(btn);
            }
        }
    }

    // 필터용 버튼들 찾기
    QList<QPushButton*> filterButtons = ui->scrollAreaWidgetContents->findChildren<QPushButton*>();

    // 필터용 버튼들에 기능 연결
    for(QPushButton* btn : filterButtons) {
        // 체크 가능한 토글 버튼인 경우에만 슬롯과 연결 (일반 버튼과 구분하기 위함)
        if(btn->isCheckable()) {
            connect(btn, &QPushButton::toggled, this, &MainWindow::onFilterButtonToggled);
        }
    }

    // 캐릭터 목록의 슬릇 클릭 시 이벤트 처리 연결
    connect(this->ui->listWidget_char, &QListWidget::itemClicked, this, &MainWindow::onCharacterClicked);

    // 스쿼드 ui 클릭시 이벤트 처리 연결
    for(auto widget : this->widget_squad){
        connect(widget, &CharacterSlotWidget::clicked, this, &MainWindow::onSquadClicked);
    }
}

void MainWindow::onFilterButtonToggled(bool checked){
    // 신호를 보낸 객체(버튼)를 찾아 QPushButton 포인터로 변환
    QPushButton *clickedButton = qobject_cast<QPushButton*>(sender());
    if (!clickedButton) return; // 버튼이 아니면 종료 (안전 장치)

    // 해당 버튼의 부모 위젯을 찾아 QGroupBox 포인터로 변환
    QGroupBox *parentGroup = qobject_cast<QGroupBox*>(clickedButton->parentWidget());
    if (!parentGroup) return; // 부모가 그룹박스가 아니면 종료

    // 텍스트 추출
    QString category_kor = parentGroup->title();
    QString category = this->getColumnName(category_kor);

    // 키워드인 경우 property에서 원본 가져오기, 아니면 text() 사용
    QString filter_value = clickedButton->property("original_keyword").toString();
    if(filter_value.isEmpty()) filter_value = clickedButton->text();

    // '증가' -> '▲', '감소' -> '▼' 치환 (DB의 기호 기반 키워드와 매칭하기 위함)
    filter_value.replace("증가", "▲");
    filter_value.replace("감소", "▼");

    if(checked){// 클릭해서 필터링 활성화시 동작
        this->map_filter[category].append(filter_value);
    }
    else{// 클릭해서 필터링 비활성화시 동작
        this->map_filter[category].removeAll(filter_value);
    }

    // 값이 존재하는 카테고리만 sql 명령어 생성
    QStringList sql_condition;

    for(auto [key, values] : this->map_filter.asKeyValueRange()){
        if(!values.isEmpty()){
            if(key == "공격" || key=="보호" || key=="특수 능력" || key=="제어" || key=="유틸리티" || key=="발동 조건"){
                // 스킬 키워드는 선택된 모든 조건을 동시에 만족(AND)하는 니케를 찾아야 함
                for(const QString &val : values){
                    QString condition = QString("id IN ("
                                            "SELECT nikke_id FROM nikke_keywords_table WHERE keyword_id IN ("
                                                "SELECT id FROM skill_keywords_table WHERE keyword_name = '%1'"
                                        "))")
                                    .arg(val);
                    sql_condition.append(condition);
                }
            }
            else{
                // 기본 속성(기업, 코드 등)은 해당 카테고리 내에서 하나라도 일치하면 됨 (OR)
                QString condition = QString("%1 IN ('%2')").arg(key, values.join("','"));
                sql_condition.append(condition);
            }
        }
    }

    // 최종 쿼리 결과에는 언제나 id가 포함되어야 이름이 중복된 니케에 대해 대응할 수 있음
    QString sql_command = QString("SELECT id,Name,Path_img_profile FROM nikke_info_table");
    if(!sql_condition.isEmpty()){
        sql_command += QString(" WHERE ") + sql_condition.join(" AND ");
    }

    this->queryAndUpdateWidget(sql_command);
}

void MainWindow::queryAndUpdateWidget(QString sql_command){
    this->ui->listWidget_char->clear();

    qDebug() << "queryAndUpdateWidget() " << sql_command;

    QSqlQuery query;
    if(query.exec(sql_command)){

        // 캐릭터 개수만큼 목록 생성
        while(query.next()){
            int char_id = query.value(0).toInt();
            QString char_name = query.value(1).toString();
            QString img_path = query.value(2).toString();

            // 리스트에 추가할 아이템 생성
            QListWidgetItem *item = new QListWidgetItem(ui->listWidget_char);

            // 커스텀 위젯 적용
            CharacterSlotWidget *widget = new CharacterSlotWidget();
            // 니케 프로필 이미지와 이름 적용
            widget->setCharacterProfileInfo(char_name, img_path);
            // db id 값을 넣어서 스쿼드에 넣을 때 쿼리를 쉽게 하기 위함.
            widget->setCharacterId(char_id);

            // 스쿼드에 편성했던 니케의 ui 상태 유지용
            for(int i=0; i<this->squad.size(); ++i){
                if(this->squad[i].db_id == widget->getCharacterId()) widget->setCharacterClicked(true);
            }

            // 리스트 위젯의 gridSize에 맞춰 아이템 크기 맞추기. 동적 생성 ui이기 때문에 해줘야함
            // 128, 256+50 => 50을 더하는 이유는 텍스트 크기까지 고려
            item->setSizeHint(widget->sizeHint());

            //캐릭터 리스트 위젯에 추가
            this->ui->listWidget_char->addItem(item);
            this->ui->listWidget_char->setItemWidget(item, widget);
        }
    }else{
        qDebug() << " | failed |";
    }
}

// 필터링 그룹박스 텍스트 -> DB column 이름
QString MainWindow::getColumnName(QString str){
    if(str == this->ui->groupBox_class->title()) return "Class";
    else if(str == this->ui->groupBox_code->title()) return "Code";
    else if(str == this->ui->groupBox_weapon->title()) return "Weapon_type";
    else if(str == this->ui->groupBox_company->title()) return "Company";
    else if(str == this->ui->groupBox_bust_step->title()) return "Bust_step";
    else if(str == this->ui->groupBox_attack->title()) return "공격";
    else if(str == this->ui->groupBox_protection->title()) return "보호";
    else if(str == this->ui->groupBox_mechanics->title()) return "특수 능력";
    else if(str == this->ui->groupBox_control->title()) return "제어";
    else if(str == this->ui->groupBox_utility->title()) return "유틸리티";
    else if(str == this->ui->groupBox_trigger->title()) return "발동 조건";
    else{
        qDebug() << "getColumnName() Error " << str;
        return "";
    }
}

void MainWindow::onCharacterClicked(QListWidgetItem *item){
    CharacterSlotWidget *widget = qobject_cast<CharacterSlotWidget*>(ui->listWidget_char->itemWidget(item));

    if(widget->getCharacterClicked()){
        //선택했던 캐릭이라면 스쿼에서 제거
        for(int i=0; i<this->squad.size(); ++i){
            if(this->squad[i].db_id == widget->getCharacterId()){
                this->squad.removeAt(i);
            }
        }
    }
    else{
        // 처음 선택하는 캐릭이라면
        // 기존 스쿼드에 없고 빈자리가 있다면 추가
        if(this->squad.size() < MAX_SQUAD_SIZE) {
            SquadMember member;
            member.db_id = widget->getCharacterId();
            member.name = widget->getCharacterName();
            member.img = widget->getCharacterImg();
            this->squad.append(member);
        }
        else{
            //자리가 꽉찼다면 아무 동작안함
            return;
        }
    }

    widget->setCharacterClicked(!widget->getCharacterClicked());

    this->updateSquadWidgets();
}

void MainWindow::onSquadClicked(CharacterSlotWidget *widget_member){
    // 스쿼드 ui에서 편성된 캐릭을 클릭하면 스쿼드에서 제거 및 캐릭터 목록에서 선택 상태해제
    if(widget_member->getCharacterId() == 0) return;

    // 스쿼드 리스트에서 해당 id 삭제
    for(int i=0; i<this->squad.size(); ++i){
        if(widget_member->getCharacterId() == this->squad[i].db_id){
            this->squad.removeAt(i);
            break;
        }
    }

    // 캐릭터 목록 UI에서 해당 캐릭터의 체크 상태 해제
    for(int i=0; i<ui->listWidget_char->count(); ++i){
        QListWidgetItem *item = ui->listWidget_char->item(i);
        CharacterSlotWidget *w = qobject_cast<CharacterSlotWidget*>(ui->listWidget_char->itemWidget(item));
        if(w && w->getCharacterId() == widget_member->getCharacterId()){
            w->setCharacterClicked(false);
            break;
        }
    }

    this->updateSquadWidgets();
}

void MainWindow::updateSquadWidgets(){
    for (int i=0; i<MAX_SQUAD_SIZE; ++i){
        if(i<this->squad.size()){
            this->widget_squad[i]->setCharacterProfileInfo(this->squad[i].name, this->squad[i].img);
            this->widget_squad[i]->setCharacterId(this->squad[i].db_id);
        }
        else{
            this->widget_squad[i]->setupInitState();
            this->widget_squad[i]->setCharacterId(0);
        }
    }
}

MainWindow::~MainWindow()
{
    delete this->ui;
}