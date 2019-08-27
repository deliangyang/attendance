import os
import time
import zipfile


if __name__ == '__main__':
    main_file = './dist/main.exe'
    while True:
        if os.path.exists(main_file):
            zip_file = zipfile.ZipFile('dist/考勤统计.zip', 'w')
            zip_file.write(main_file, '考勤统计.exe')
            zip_file.close()
            os.unlink(main_file)
            break
        time.sleep(2)

