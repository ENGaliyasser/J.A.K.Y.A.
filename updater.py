# Author: Ali Yasser Ali Abdallah
# Date: 4 August 2024
# Description: This file provides a modular auto-updating mechanism for PyQt applications.
#              It checks for the latest release from a specified GitHub repository, compares
#              it with the current version of the application, and if a newer version is found,
#              downloads and installs the update. The update process is managed in a separate
#              thread, with a progress bar to indicate the download progress. Additionally,
#              it can delete old versions of the application to save space.

import os
import sys
import requests
from PyQt5 import QtCore, QtWidgets


class DownloadThread(QtCore.QThread):
    """
    A QThread subclass that handles the downloading of the latest version of the application.

    Signals:
        progress (int): Emitted with the current download progress percentage.
        finished (str): Emitted when the download is completed or fails, with a message describing the outcome.
    """

    progress = QtCore.pyqtSignal(int)
    finished = QtCore.pyqtSignal(str)

    def __init__(self, download_url, save_path):
        """
        Initializes the download thread with the URL of the file to be downloaded and the path to save it.

        Args:
            download_url (str): The URL of the file to be downloaded.
            save_path (str): The local file path where the downloaded file will be saved.
        """
        super().__init__()
        self.download_url = download_url
        self.save_path = save_path

    def run(self):
        """
        Starts the download process. Emits the progress signal as the download proceeds,
        and emits the finished signal when the download is complete or fails.
        """
        try:
            response = requests.get(self.download_url, stream=True)
            total_length = response.headers.get('content-length')

            if total_length is None:  # no content length header
                self.finished.emit("Download Failed: No content length header")
                return

            total_length = int(total_length)
            downloaded = 0

            with open(self.save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=4096):
                    if chunk:  # Filter out keep-alive new chunks
                        f.write(chunk)
                        downloaded += len(chunk)
                        progress = int(100 * downloaded / total_length)
                        self.progress.emit(progress)

            self.finished.emit("Download Completed")

        except requests.exceptions.RequestException as e:
            self.finished.emit(f"Download Failed: {str(e)}")


class Updater:
    """
    The Updater class checks for updates, manages the download process, and handles
    the installation of new versions of the application. It can also delete older
    versions of the application to free up space.
    """

    def __init__(self, current_version, repo_owner, repo_name, progress_bar=None):
        """
        Initializes the Updater with the current application version, GitHub repository details,
        and an optional progress bar for visual feedback.

        Args:
            current_version (str): The current version of the application.
            repo_owner (str): The owner of the GitHub repository.
            repo_name (str): The name of the GitHub repository.
            progress_bar (QtWidgets.QProgressBar, optional): The progress bar widget to show download progress.
        """
        self.current_version = current_version
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.progress_bar = progress_bar

    def check_update(self, update_button, ask_label):
        """
        Checks the GitHub repository for the latest release. If a new version is available,
        it makes the update button and ask label visible to the user.

        Args:
            update_button (QtWidgets.QPushButton): The button that allows the user to trigger the update.
            ask_label (QtWidgets.QLabel): The label asking the user if they want to update.
        """
        api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"

        try:
            response = requests.get(api_url)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release['tag_name']
            self.download_url = latest_release['assets'][0]['browser_download_url']

            if self.current_version != latest_version:
                update_button.setVisible(True)
                ask_label.setVisible(True)
                QtWidgets.QMessageBox.information(None, "Update Available",
                                                  f"New version {latest_version} is available!")
                self.latest_version = latest_version
            else:
                QtWidgets.QMessageBox.information(None, "No Update", "You already have the latest version.")

        except requests.exceptions.RequestException as e:
            QtWidgets.QMessageBox.critical(None, "Error", f"Failed to check for updates: {str(e)}")

    def update_application(self):
        """
        Initiates the update process by starting the download of the new version.
        The download progress is shown on the specified progress bar.
        """
        if not hasattr(self, 'download_url') or not hasattr(self, 'latest_version'):
            QtWidgets.QMessageBox.critical(None, "Error", "No update URL or version found.")
            return

        current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        new_exe_path = os.path.join(current_dir, f"JAKYA version_{self.latest_version}.exe")

        # Show the progress bar and start the download in a separate thread
        if self.progress_bar:
            self.progress_bar.setVisible(True)

        self.download_thread = DownloadThread(self.download_url, new_exe_path)
        self.download_thread.progress.connect(self.progress_bar.setValue)
        self.download_thread.finished.connect(lambda msg: self.on_download_finished(msg, new_exe_path))
        self.download_thread.start()

    def on_download_finished(self, msg, new_exe_path):
        """
        Handles the completion of the download process. If the download was successful,
        the current version is renamed, and the new version is launched. The current
        version exits after the new version is started.

        Args:
            msg (str): The message indicating the result of the download process.
            new_exe_path (str): The file path of the downloaded executable.
        """
        if "Completed" in msg:
            QtWidgets.QMessageBox.information(None, "Download Completed", "Update downloaded successfully.")

            # Rename the current executable to "tobedeletedddddd.exe"
            current_executable = os.path.abspath(sys.argv[0])
            new_name = os.path.join(os.path.dirname(current_executable), "tobedeletedddddd.exe")
            os.rename(current_executable, new_name)

            # Start the new version of the application
            os.startfile(new_exe_path)
            sys.exit(0)  # Exit the current application

        else:
            QtWidgets.QMessageBox.critical(None, "Error", msg)

        if self.progress_bar:
            self.progress_bar.setVisible(False)

    @staticmethod
    def delete_old_versions(current_version):
        """
        Deletes the "tobedeletedddddd.exe" file from the current directory if it exists.

        Args:
            current_version (str): The current version of the application, used to determine which files to keep.
        """
        current_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        to_delete = os.path.join(current_dir, "tobedeletedddddd.exe")

        if os.path.exists(to_delete):
            try:
                os.remove(to_delete)
            except Exception as e:
                QtWidgets.QMessageBox.critical(None, "Error", f"Failed to delete {to_delete}.\nError: {str(e)}")
