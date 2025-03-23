import ftplib
import os
import sys
import log_handler


class FtpHandler:

    def downloadFiles(self, ftp, path, destination):
        org_cwd = os.getcwd()
        # path & destination are str of the form "/dir/folder/something/"
        # path should be the abs path to the root FOLDER of the file tree to download
        try:
            ftp.cwd(path)
            ftp.retrlines('LIST')
            # clone path to destination
            if "/" in path:
                path = "/" + path.split("/")[-2]
            print(path)
            os.makedirs(destination + path, exist_ok=True)

            os.chdir(destination)

            print(destination + path + " built")

        except ftplib.error_perm:
            # invalid entry (ensure input form: "/dir/folder/something/")
            print(
                "error: could not change to " + path)
            sys.exit("ending session")

        # list children:
        filelist = ftp.nlst()
        print(filelist)
        for file in filelist:
            try:
                #print(path + file + "/")
                # this will check if file is folder:
                print("HER " + file + "/")
                ftp.cwd(file + "/")
                # if so, explore it:
                print("k")
                self.downloadFiles(ftp, file + "/", destination + path)
            except ftplib.error_perm:
                # not a folder with accessible content
                # download & return
                print(os.path.join(destination, file))
                os.makedirs(destination, exist_ok=True)
                os.chdir(destination + path)
                # possibly need a permission exception catch:

                with open(os.path.join(destination + path, file.split("/")[-1]), "wb") as f:
                    ftp.retrbinary("RETR " + file, f.write)
                print(file + " downloaded")
        os.chdir(org_cwd)
        return

    def download_newest_zelda_save_from_switch(self):
        # ftp = ftplib.FTP("192.168.1.207", user="Lund", passwd="Lund1234")
        # ftp.login("Lund", "Lund1234")

        ftp = ftplib.FTP()
        ftp.connect('192.168.1.207', 5000)
        ftp.login('Lund', 'Lund1234')

        ftp.cwd('/switch/Checkpoint/saves/0x0100F2C0115B6000 The Legend of Zelda  Tears of the Kingdom/')
        #ftp.retrlines('LIST')
        content = []

        ftp.dir(content.append)

        dates = [date for line in content for date in line.split(" ") if "-" in date]

        #ftp.cwd("/")

        #print(content)
        #print(dates[-1])

        #TODO check if switch version is older than yuzu save

        self.downloadFiles(ftp, dates[-1] + " Jakob Lund", "C:/Users/Jakob/Desktop/zelda_saves/")

        ftp.quit()
