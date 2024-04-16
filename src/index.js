const fs = require('fs');
const {Workbook} = require('exceljs')
const data = require('./data');

const apiToken = 'SET_API_TOKEN_HERE';
const getAccessTokenUrl = (clientId, clientSecret, code) => `https://oauth.vk.com/access_token?client_id=${clientId}&client_secret=${clientSecret}&redirect_uri=http://localhost&code=${code}`;
const getGroupsByIdsUrl = 'https://api.vk.com/method/groups.getById';

const getGroupsByIds = async (accessToken, groupsIds) => {
    const formData = new FormData();
    formData.set('access_token', accessToken);
    formData.set('group_ids', groupsIds.join(','));
    formData.set('v', '5.199');

    return fetch(getGroupsByIdsUrl, {
        method: 'post',
        body: formData,
    })
}

const groupsInfo = {};
const requests = [];
for (let i = 0; i <= data.length; i += 400) {
    requests.push(
        getGroupsByIds(apiToken, data.slice(i, i + 400))
            .then((response) => response.json())
            .then((data) => data.response.groups)
            .then((groups) => groups.forEach((group) => groupsInfo[group.id] = group.name))
    )
}

let workBook;
let workSheet;

Promise
    .all(requests)
    .then(() => workBook = new Workbook())
    .then(() => workSheet = workBook.addWorksheet('sheet'))
    .then(() => {
        workSheet.columns = [
            {header: 'Club ID', key: 'id', width: 10,},
            {header: 'Club Name', key: 'name', width: 30,}
        ];
    })
    .then(() => {
        workSheet.getColumn('id').values = Object.keys(groupsInfo);
        workSheet.getColumn('name').values = Object.values(groupsInfo);
    })
    .then(() => workBook.xlsx.writeFile('names2.xlsx'))
