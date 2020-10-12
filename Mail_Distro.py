from Mail_Distro_Class import MailDistro
from Mail_Distro_Settings import local_config


if __name__ == '__main__':
    if local_config['Distro_Configs']:
        obj = MailDistro()

        for distro, table in local_config['Distro_Configs'].items():
            obj.process(distro, table)
