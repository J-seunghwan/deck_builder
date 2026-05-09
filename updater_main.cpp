#include <QCoreApplication>
#include <QProcess>
#include <QDir>
#include <QFile>
#include <QThread>
#include <QDebug>
#include <QStandardPaths>

/**
 * Updater Program
 * 1. Wait for the main application (Deck_Builder.exe) to terminate.
 * 2. Unzip the downloaded update file using PowerShell.
 * 3. Restart the main application.
 */
int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    if (argc < 2) {
        qDebug() << "Usage: updater.exe <update_zip_path>";
        return -1;
    }

    QString zipPath = argv[1];
    QString appPath = QCoreApplication::applicationDirPath();
    QString mainExe = appPath + "/Deck_Builder.exe";

    // 1. 메인 프로그램이 완전히 종료될 때까지 대기 (최대 10초)
    int retry = 0;
    while (QFile::exists(mainExe) && !QFile::remove(mainExe)) {
        if (retry++ > 10) break; 
        QThread::sleep(1);
    }

    // 2. PowerShell을 이용하여 압축 해제
    // -Force 옵션을 사용하여 기존 파일 덮어쓰기
    QStringList arguments;
    arguments << "-NoProfile" << "-Command" 
              << QString("Expand-Archive -Path '%1' -DestinationPath '%2' -Force").arg(zipPath, appPath);

    QProcess process;
    process.start("powershell.exe", arguments);
    process.waitForFinished(60000); // 최대 1분 대기

    if (process.exitCode() == 0) {
        qDebug() << "Update successful.";
        // 임시 파일 삭제 시도
        QFile::remove(zipPath);
    } else {
        qDebug() << "Update failed:" << process.readAllStandardError();
    }

    // 3. 메인 프로그램 재실행
    QProcess::startDetached(mainExe);

    return 0;
}
