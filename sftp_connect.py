import paramiko
import os
import sftp_info

class Sftp:
    HOST = sftp_info.HOST
    PORT = sftp_info.PORT
    USERNAME = sftp_info.USERNAME
    PW = sftp_info.PW

    local_from_path = None
    local_to_path = None
    remote_from_path = None

    sftp = None
    d = None

    def __init__(self,env,d,local_from_folder,local_to_folder):
        transprot = paramiko.transport.Transport(self.HOST,self.PORT)
        transprot.connect(username=self.USERNAME, password=self.PW)
        self.sftp = paramiko.SFTPClient.from_transport(transprot)

        self.local_from_path = local_from_folder + "/" + d
        self.local_to_path = local_to_folder + "/" + d
        self.remote_from_path = env + sftp_info.BANK_PATH + d
        self.remote_to_path = env + sftp_info.BANK_PATH + d + "/pdf"

    def get_file_from_sftp(self):
        '''
        remote_from_folder -> local_from_folder로 파일을 다운로드한다
        '''
        sftp_files = self.sftp.listdir(self.remote_from_path)

        if not os.path.exists(self.local_from_path):
            os.mkdir(self.local_from_path)

        for s_f in sftp_files:
            try:
                print(self.remote_from_path+"/"+s_f, "->", self.local_from_path+"/"+s_f)
                self.sftp.get(self.remote_from_path+"/"+s_f, self.local_from_path+"/"+s_f)
            except:
                pass


    def put_file_to_sftp(self):
        '''
        local_to_folder -> remote_to_folder(PDF)로 파일을 업로드한다
        '''
        local_files = os.listdir(self.local_to_path)

        try:
            self.sftp.mkdir(self.remote_to_path)
        except:
            pass

        for l_f in local_files:
            print(self.local_to_path+"/"+l_f, "->" ,self.remote_to_path +"/"+l_f)
            self.sftp.put(self.local_to_path+"/"+l_f, self.remote_to_path +"/"+l_f)

        self.sftp.close()
