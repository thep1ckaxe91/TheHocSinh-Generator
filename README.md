### Installation
- `git pull`
- Put your excel data in `excels_files` folder
- Put your `default.png` into `assets`
- Adjust the const for each excel files (most of the time) and default image (literally change per year) in `export-config.yml`

### TODO:
- Make CONST adjustable (or automatically someday) on files instead of source code for build run
- Make it Cross Platform (probably did)
- Performance improvement possibility since windows doesnt need to be display, and save task can be divided into many threads by using many surface and PIL image saving