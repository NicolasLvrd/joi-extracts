const {MongoClient} = require('mongodb');
const XLSX = require('xlsx');
const {DB_URI} = require('./db');
//const fs = require("fs");

async function main(){

    const uri = DB_URI;
 
    const joi_id = 'joi-2022';

    const client = new MongoClient(uri);

    const db = client.db("forumorg-prod");

    const activities = {};

    const activities_titles = {};
 
    try {
        // connection au cluster
        await client.connect();
        
        // initialisation d'un dictionnaire type { activity_id: { slot_id: [student #1, student #2] } }
        const joi = await db.collection("events").findOne( {'id': 'joi-2022'} );
        joi.round_tables.forEach( (round_table) => { // pour chaque table ronde
            slot_dict = {};
            for (slot in round_table.slots) {
                slot_dict[slot] = [];
            }
            activities[round_table.id] = slot_dict;
            activities_titles[round_table.id] = round_table.title; // dictionnaire liant l'id de l'activité à son titre
        })
        joi.conferences.forEach( (conference) => { // pour chaque conférence
            slot_dict = {};
            for (slot in conference.slots) {
                slot_dict[slot] = [];
            }
            activities[conference.id] = slot_dict;
            activities_titles[conference.id] = conference.title; // dictionnaire liant l'id de l'activité à son titre
        })
        
        // recupération des étudiants inscrits à l'événement et ajout dans le dictionnaire
        await db.collection("users").find( {'events.joi-2022.slots': {$exists: true}} )
        .forEach( user => {
            for (slot in user.events[joi_id].slots) {
                //if (user.events[joi_id].slots[slot] in ('k9avr', 'gFeuu', '0PPv3', 'ce3UP', 'YgLau', 'lsXwb', 'eAZAp', 'rfWTz', '36VsS', 'wAqBY', 'nUCfG', 'MYRIR', 'XFNDe', '6uDOG')) {
                    activities[ user.events[joi_id].slots[slot] ][ slot ].push([user.profile.first_name, user.profile.name, user.id])
                //}
            }
        })

        // création du fichier excel
        for (activity_id in activities) {
            for (slot_id in activities[activity_id]) {
                /* écrire un CSV
                var csv = "";
                csv += activity_id + "\r\n";
                csv += slot_id + "\r\n";
                csv += header.join(";") + "\r\n";
                for (let i of activities[activity_id][slot_id]) {
                    csv += i.join(";") + "\r\n";
                }
                fs.writeFileSync(activity_id + '_' + slot_id.split(' - ')[0] + '_' + slot_id.split(' - ')[1] +'.csv', csv);
                */

                // écrire un XLSX
                const title = [activities_titles[activity_id]];
                const slot = [slot_id];
                const header = ['first_name', 'name', 'email'];
                const data = activities[activity_id][slot_id];
                data.unshift(header);
                data.unshift(['']);
                data.unshift(slot);
                data.unshift(title);


                const merge1 = XLSX.utils.decode_range("A1:C1");
                const merge2 = XLSX.utils.decode_range("A2:C2");
                
                const workSheet = XLSX.utils.aoa_to_sheet(data);
                
                if(!workSheet['!merges']) workSheet['!merges'] = [];
                workSheet['!merges'].push(merge1);
                workSheet['!merges'].push(merge2);
                
                const workBook = XLSX.utils.book_new();XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet 1");
                XLSX.writeFile(workBook, './output/'+activity_id + '_' + slot_id.split(' - ')[0] + '_' + slot_id.split(' - ')[1]+".xlsx");
            }
        }

        console.log("done");

    } catch (e) {
        console.error(e);
    } finally {
        await client.close();
    }
}

main().catch(console.error);