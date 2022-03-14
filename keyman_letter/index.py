from utils import libs

def main():
    libs.setVariables(
        max_delay_time=20,  # No need to change this
        total_page=2,  # Need to change based on the number of pagination
        list_classname='list',
        rows_classname='jss168',
        minimize_windows=False,
        driver_path='../drivers/chromedriver.exe',  # No need to change this path 97.0.42
        docker=False,  # Change to True if you want to use docker
    )
    libs.scrapFirstPage()
    libs.scrapRemainingPage()
    results = libs.getResults()
    libs.exportExcel(results)

if __name__ == "__main__":
    main()
