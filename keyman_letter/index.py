from keyman_letter.utils import libs

def main():
    libs.setVariables(
        max_delay_time=20,
        total_page=2,
        list_classname='list',
        rows_classname='jss166',
        driver_path='../drivers/chromedriver.exe',
        minimize_windows=True
    )
    libs.scrapFirstPage()
    libs.scrapRemainingPage()
    results = libs.getResults()
    libs.exportExcel(results)

if __name__ == "__main__":
    main()
