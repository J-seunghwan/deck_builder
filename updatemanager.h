#ifndef UPDATEMANAGER_H
#define UPDATEMANAGER_H

#include <QObject>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QFile>
#include <QTimer>

class UpdateManager : public QObject
{
    Q_OBJECT
public:
    explicit UpdateManager(QObject *parent = nullptr);
    void startUpdateCheck(); // 시작점 (재시도 포함)

signals:
    void updateAvailable(const QString &version, const QString &downloadUrl);
    void noUpdateFound();
    void downloadProgress(qint64 bytesReceived, qint64 bytesTotal);
    void downloadFinished(const QString &filePath);
    void errorOccurred(const QString &error);

private slots:
    void checkForUpdates();
    void onReleaseInfoReceived(QNetworkReply *reply);
    void downloadUpdate(const QString &url);
    void onDownloadFinished(QNetworkReply *reply);

private:
    bool isNewerVersion(const QString &latest, const QString &current);
    QNetworkAccessManager *networkManager;
    QTimer *retryTimer;
    QString currentVersion;
    QString repoPath = "J-seunghwan/deck_builder";
    bool hasCheckedSuccessfully = false; // 성공적으로 체크했는지 여부
};

#endif // UPDATEMANAGER_H
