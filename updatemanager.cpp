#include "updatemanager.h"
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QStandardPaths>
#include <QCoreApplication>
#include <QDir>

UpdateManager::UpdateManager(QObject *parent) : QObject(parent)
{
    networkManager = new QNetworkAccessManager(this);
    currentVersion = APP_VERSION;
    
    retryTimer = new QTimer(this);
    retryTimer->setInterval(10000); // 10초마다 재시도
    connect(retryTimer, &QTimer::timeout, this, &UpdateManager::checkForUpdates);
}

void UpdateManager::startUpdateCheck()
{
    if (hasCheckedSuccessfully) return;
    checkForUpdates();
}

void UpdateManager::checkForUpdates()
{
    if (hasCheckedSuccessfully) {
        retryTimer->stop();
        return;
    }

    QUrl url(QString("https://api.github.com/repos/%1/releases/latest").arg(repoPath));
    QNetworkRequest request(url);
    request.setHeader(QNetworkRequest::UserAgentHeader, "Qt6-UpdateManager");
    
    QNetworkReply *reply = networkManager->get(request);
    connect(reply, &QNetworkReply::finished, this, [this, reply]() {
        onReleaseInfoReceived(reply);
    });
}

void UpdateManager::onReleaseInfoReceived(QNetworkReply *reply)
{
    reply->deleteLater();
    
    // 네트워크 오류 시 재시도 타이머 시작
    if (reply->error() != QNetworkReply::NoError) {
        if (!retryTimer->isActive()) {
            retryTimer->start();
        }
        emit errorOccurred("Connection failed, retrying in 10s... " + reply->errorString());
        return;
    }

    // 성공적으로 응답을 받음
    hasCheckedSuccessfully = true;
    retryTimer->stop();

    QByteArray data = reply->readAll();
    QJsonDocument doc = QJsonDocument::fromJson(data);
    QJsonObject json = doc.object();

    QString latestVersion = json.value("tag_name").toString();
    if (latestVersion.isEmpty()) {
        emit noUpdateFound();
        return;
    }

    if (isNewerVersion(latestVersion, currentVersion)) {
        QJsonArray assets = json.value("assets").toArray();
        QString downloadUrl;
        for (const QJsonValue &asset : assets) {
            QString name = asset.toObject().value("name").toString();
            if (name.endsWith(".zip")) {
                downloadUrl = asset.toObject().value("browser_download_url").toString();
                break;
            }
        }

        if (!downloadUrl.isEmpty()) {
            emit updateAvailable(latestVersion, downloadUrl);
            downloadUpdate(downloadUrl); // 자동 다운로드 시작
        } else {
            emit errorOccurred("No suitable download asset found.");
        }
    } else {
        emit noUpdateFound();
    }
}

void UpdateManager::downloadUpdate(const QString &url)
{
    QNetworkRequest request((QUrl(url)));
    request.setAttribute(QNetworkRequest::RedirectPolicyAttribute, QNetworkRequest::NoLessSafeRedirectPolicy);
    
    QNetworkReply *reply = networkManager->get(request);
    
    connect(reply, &QNetworkReply::downloadProgress, this, &UpdateManager::downloadProgress);
    connect(reply, &QNetworkReply::finished, this, [this, reply]() {
        onDownloadFinished(reply);
    });
}

void UpdateManager::onDownloadFinished(QNetworkReply *reply)
{
    reply->deleteLater();
    if (reply->error() != QNetworkReply::NoError) {
        emit errorOccurred("Download failed: " + reply->errorString());
        return;
    }

    QString tempPath = QDir::tempPath() + "/update.zip";
    QFile file(tempPath);
    if (file.open(QIODevice::WriteOnly)) {
        file.write(reply->readAll());
        file.close();
        emit downloadFinished(tempPath);
    } else {
        emit errorOccurred("Failed to save temporary update file.");
    }
}

bool UpdateManager::isNewerVersion(const QString &latest, const QString &current)
{
    // 'v1.0.0' 같은 형식에서 'v' 제거
    auto normalize = [](QString v) {
        if (v.startsWith('v', Qt::CaseInsensitive)) v.remove(0, 1);
        return v.split('.');
    };

    QStringList latestParts = normalize(latest);
    QStringList currentParts = normalize(current);

    int maxLen = std::max(latestParts.size(), currentParts.size());
    for (int i = 0; i < maxLen; ++i) {
        int l = (i < latestParts.size()) ? latestParts[i].toInt() : 0;
        int c = (i < currentParts.size()) ? currentParts[i].toInt() : 0;

        if (l > c) return true;
        if (l < c) return false;
    }

    return false;
}
