# A2A Data Scrapping
- [Overview](#overview)
- [Requirements](#requirements)
- [Installation](#steps-to-run-this-on-your-local)
  - [1. **Clone the application**](#1-clone-the-application)
  - [2. **Install necessary dependencies for the application**](#2-install-necessary-dependencies-for-the-application)
  - [3. **Create a .env file and copy the contents from .env.example**](#3-create-a-env-file-and-copy-the-contents-from-envexample)
  - [4. **Start the application**](#4-start-the-application)

## Overview
This is a clone application for trello. This has been built for learning purpose. My plan is to improve this project and add more features in every release.

## Requirements
1. [Node.js](https://nodejs.org/)
2. [npm](https://www.npmjs.com/)

## Steps to run this on your local
First install the MongoDB Compass for better visualization of data with MongoDB server.

1. Clone this repo using `git clone https://github.com/knowankit/trello-clone.git`
2. Create _.env.local_ and add this env variable `LOCAL_MONGODB=mongodb://localhost:27017/trello`
    Add `JWT_SECRET_KEY=randomstrings`
3. Run `yarn install`
4. Run `yarn dev`

`For unsplash gallery, api key is needed which can be generated from unsplash website`
