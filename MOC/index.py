from utils import _libs

def main():
    _libs.setVariables(start_id_at=0, stop_id_at=3, driver_path='../../drivers/chromedriver.exe',)
    _libs.run()

if __name__ == "__main__":
    main()
