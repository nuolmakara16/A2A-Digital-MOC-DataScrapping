from keyman_letter.utils import libs

def main():
    libs.setVariables(
        max_delay_time=20,  # No need to change this
        total_page=335,  # Need to change based on the number of pagination
        list_classname='list',
        rows_classname='jss166',
        driver_path='../drivers/chromedriver.exe',  # No need to change this path 97.0.42
        minimize_windows=False
    )
    libs.scrapFirstPage()
    libs.scrapRemainingPage()
    results = libs.getResults()
    libs.exportExcel(results)

if __name__ == "__main__":
    main()
