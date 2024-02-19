/*
    參考:
        https://b-l-u-e-b-e-r-r-y.github.io/post/PTTCrawler/
*/

const fs = require("fs");
const request = require("request");
const cheerio = require("cheerio");
const ExcelJS = require('exceljs');

/*
    解析 播放資訊
    0: 首播日, 1: 播放星期, 2: 播放時間
*/
const playInf_processing = (playInf) => {
    return new Promise((resolve, reject) => {
        let crossingFlag = false;
        let premiereExist;

        const WeeklyIndex = {
            '日': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
        }

        playInf = playInf.split('／');  // 字串分割
        premiereExist = (playInf.length == 3)? true : false;

        playInf = (premiereExist)? {
            premiere: playInf[0],
            playWeek: playInf[1],
            playTime: playInf[2]
        } : {
            premiere: '',
            playWeek: playInf[0],
            playTime: playInf[1]
        };

        // playInf.premiere = playInf.premiere.replace('起', '');
        if(premiereExist){
            let playInf_tmp = playInf.premiere.match(/(\d+)月(\d+)日起/);
            playInf.premiere = {
                m: parseInt(playInf_tmp[1], 10),
                d: parseInt(playInf_tmp[2], 10),
            };
        }


        playInf.playWeek = playInf.playWeek.replace('每週', '');
        /*  判斷是否跨日 */
        if(playInf.playWeek.includes('深夜')){
            playInf.playWeek = playInf.playWeek.replace('深夜', '');
            crossingFlag = true;
        }
        playInf.playWeek = WeeklyIndex[playInf.playWeek];

        playInf_tmp = playInf.playTime.match(/(\d{1,2})時(\d{1,2})分/);
        playInf.playTime = {
            h: parseInt(playInf_tmp[1], 10),
            m: parseInt(playInf_tmp[2], 10),
        };

        /* 跨日處裡 */
        if(crossingFlag){
            if(premiereExist) playInf.premiere.d++;
            playInf.playWeek = (playInf.playWeek == 6)? 0 : playInf.playWeek + 1;
            playInf.playTime.h -= 24;
        }

        if(premiereExist) playInf.premiere = `2024-${playInf.premiere.m}-${playInf.premiere.d}` 
        playInf.playTime = `${playInf.playTime.h}:${playInf.playTime.m}`

        resolve(playInf);
    });
};

/*
    解析 音樂
    取得OP ED
*/
const animeMusic_processing = (animeMusic) => {
    return new Promise((resolve, reject) => {
        let animeMusiclist = {};

        if(animeMusic == '') resolve('');

        animeMusic.each((index) => {
            let musicType = animeMusic.eq(index).find('.song_type').text();
            
            animeMusiclist[musicType] = {
                playLink: animeMusic.eq(index).find('a.youtube').attr('href'),
            };
        });

        resolve(JSON.stringify(animeMusiclist));
    })
};

/*
    解析 播放平台
*/
const animeStreams_processing = (animeStreams) => {
    return new Promise((resolve, reject) => {
        let TWindex = 0;
        let streamArea = animeStreams.find('.stream-area');
        let platlist = {};

        streamArea.each((index) => {
            if(streamArea.eq(index).text() == '台灣'){
                TWindex = index;
            }
        });

        if(TWindex == 0){
            resolve('');
        }

        let platforms = animeStreams.eq(TWindex).find('.steam-site-item');
        
        platforms.each((index) => {
            let platform = platforms.eq(index); // 播放平台

            let playTime = platform.find('.steam-site-time.time_today').text();
            let playTimeStr = playTime.match(/\d{1,2}:\d{2}/);
            playTime = (playTimeStr)? playTimeStr[0] : playTime;

            // 播放時間處裡
            if(playTime != ''){
                playTime = playTime.split(':');
                playTime[0] = parseInt(playTime[0], 10);
                playTime[0] = (playTime[0] >= 24)? playTime[0] - 24 : playTime[0]; // 跨日處裡
                playTime  = `${playTime[0]}:${playTime[1]}`;
            }

            platlist[platform.find('.steam-site-name').text()] = {
                platformLink: platform.find('a.stream-site').attr('href'),
                playTime    : playTime,
            }
        });

        // fs.writeFileSync("JSON_output.json", JSON.stringify(platlist), "utf-8");

        resolve(JSON.stringify(platlist));
    })
};

/*
    解析 外部連結
    目前只取得 官方網站
*/
const externalLink_processing = (externalLink) => {
    return new Promise((resolve, reject) => {
        let official_web_index = -1;

        externalLink.each((index) => {
            if(externalLink.eq(index).text() == '官方網站'){
                official_web_index = index;
            }
        })

        if(official_web_index == -1) resolve('');

        let officialWeb = externalLink.eq(official_web_index).attr('href');

        resolve(officialWeb);
    })
};

/*
    解析 製作陣容
*/
const staff_processing = (staff) => {
    return new Promise((resolve, reject) => {
        let staffList = {};

        staff.each((index) => {
            let person = staff.eq(index).text();
            let person_tmp = person.split('：');
            staffList[person_tmp[0]] = person_tmp[1]
        });

        fs.writeFileSync("JSON_output.json", JSON.stringify(staffList), "utf-8");

        resolve(JSON.stringify(staffList));
    })
};

/*
    解析 預告PV
*/
const pv_processing = (pvs) => {
    return new Promise((resolve, reject) => {
        let pv = [];

        if(pvs == '') resolve('');

        pvs.each((index) => {
            pv.push(pvs.eq(index).attr('href'));
        });

        resolve(JSON.stringify(pv));
    })
};

const crawler = (url) => {
    return new Promise((resolve, reject) => {
        request({
            url: url,
            method: "GET"
        }, async(error, res, body) => {
            // 如果有錯誤訊息，或沒有 body(內容)，就 return
            if (error || !body) {
                console.log('error!');
                resolve();
            }

            // 使用 cheerio 解析 HTML 內容
            const $ = cheerio.load(body);

            let worksData_out = [];
            let worksData_in = $('.clear-both.acgs-anime-block');

            for(let work of worksData_in){
                let animeName = $(work).find('.anime_info.main.site-content-float .entity_localized_name').text();              // 動漫名稱
                let animeImgURL = $(work).find('.overflow-hidden.anime_cover_image .img-fit-cover').attr('acgs-img-data-url');  // 圖片
                let playInf = $(work).find('.time_today.main_time').text(); // 播放資訊
                let animeStory = $(work).find('.anime_story').text();       // 故事大綱
                let animeMusic = $(work).find('.anime_music');              // 動漫音樂
                let animeStreams = $(work).find('.anime_streams');          // 播放平台
                let externalLink = $(work).find('.anime_links a.normal');   // 外部連結(官網)
                let staff = $(work).find('.anime_staff .anime_person');     // 製作陣容
                let pv = $(work).find('.anime_trailers a.youtube');         // 預告片

                [
                    playInf,
                    animeMusic,
                    animeStreams,
                    externalLink,
                    staff,
                    pv,
                ] = await Promise.all([
                    playInf_processing(playInf),
                    animeMusic_processing($(animeMusic)),
                    animeStreams_processing($(animeStreams)),
                    externalLink_processing($(externalLink)),
                    staff_processing($(staff)),
                    pv_processing($(pv)),
                ])

                worksData_out.push({
                    animeName       : animeName,
                    animeImgURL     : animeImgURL,
                    ...playInf,
                    animeStory      : animeStory,
                    animeMusic      : animeMusic,
                    animeStreams    : animeStreams,
                    officialWeb     : externalLink,
                    staff           : staff,
                    pv              :pv,
                });
            }

            // let work = $(worksData_in[52]).html();
            // let animeName = $(work).find('.anime_info.main.site-content-float .entity_localized_name').text();              // 動漫名稱
            // let animeImgURL = $(work).find('.overflow-hidden.anime_cover_image .img-fit-cover').attr('acgs-img-data-url');  // 圖片
            // let playInf = $(work).find('.time_today.main_time').text(); // 播放資訊
            // let animeStory = $(work).find('.anime_story').text();       // 故事大綱
            // let animeMusic = $(work).find('.anime_music');
            // let animeStreams = $(work).find('.anime_streams');
            // let externalLink = $(work).find('.anime_links a.normal');
            // let staff = $(work).find('.anime_staff .anime_person');
            // let pv = $(work).find('.anime_trailers a.youtube');

            // playInf = await playInf_processing(playInf);
            // animeMusic = await animeMusic_processing($(animeMusic));
            // animeStreams = await animeStreams_processing($(animeStreams));
            // externalLink = await externalLink_processing($(externalLink));
            // staff = await staff_processing($(staff));
            // pv = await pv_processing($(pv));

            // worksData_out.push({
            //     animeName       : animeName,
            //     animeImgURL     : animeImgURL,
            //     ...playInf,
            //     animeStory      : animeStory,
            //     animeMusic      : animeMusic,
            //     animeStreams    : animeStreams,
            //     officialWeb     : externalLink,
            //     staff           : staff,
            //     pv              :pv
            // });

            // console.log(worksData_out);

            resolve(worksData_out);
        });
    });
};

const excel_convert_save = (worksData) => {
    const getAttrName = (keys) => {
        return new Promise((resolve, reject) => {
            let columns = [];
    
            for(let keyName of keys){
                columns.push({
                    header: keyName,
                    key: keyName
                });
            }
    
            resolve(columns);
        });
    };

    return new Promise(async (resolve, reject) => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('worksData 1');
        const keys = Object.keys(worksData[0]);

        worksheet.columns = await getAttrName(keys);    // 獲取屬性(鍵值)

        worksheet.addRows(worksData);

        workbook.xlsx.writeFile('animeData.xlsx')
        .then(() => {
            console.log('Excel文件創建成功');
            resolve(true);
        })
        .catch((err) => {
            console.error('創建Excel時發生錯誤：', err);
            resolve(false);
        });
    });
};

(async() => {
    let worksData = await crawler('https://acgsecrets.hk/bangumi/202401/');
    await excel_convert_save(worksData);

    // console.log(worksData, `\n\n共${worksData.length}部作品`);
})();
