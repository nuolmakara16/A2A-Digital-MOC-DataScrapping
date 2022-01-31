# A2A Data Scrapping
- [Overview](#overview)
- [Requirements](#requirements)
- [Installation](#steps-to-run-this-on-your-local)

## Overview
The idea of this project is to govern or gather information and data from various website to serve marketing purpose

## Requirements
1. Clone this repo using `git clone https://github.com/nuolmakara16/A2A-Digital-MOC-DataScrapping.git`
2. [Anaconda](https://www.anaconda.com/products/individual)
3. [Google chrome driver](https://chromedriver.chromium.org/downloads) or [Docker](https://www.docker.com/products/docker-desktop)

### You can skip google chrome driver if you have docker install

```bash
$ docker-compose up -d --build
```
- By default, the project will run without docker, so you have to download the chrome driver based on your Chrome version and put it inside drivers folder
## Setup python environment
1. Create new conda environment `conda create --name py36 python=3.6`
2. Activate env `conda activate py36`
3. Navigate to the project you want to run `cd MOC || cd keyman_letter`
4. Install dependencies `pip install -r requirements.txt`
5. Run the project `python index.py`

### How to run Keyman Letters Project
````pycon
    libs.setVariables(
        # Change docker option to True if you use docker-compose up
        docker=True,  # Change to True if want to use docker
    )
````
```bash
$ cd keyman_letter
$ python index.py
```

### How to run Keyman Letters Project
```bash
$ cd MOC
$ python index.py
```