import { useState, useEffect, useCallback, useRef } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from "recharts";
import * as XLSX from 'xlsx';

const LOGO = "https://cdn.discordapp.com/attachments/719554714093617172/992416002992255056/Blue_logo2x.png?ex=6a012bad&is=69ffda2d&hm=832e3e89daf693d926475526d6a2822e4c3b5a4eda56b01e9cc2cd3f4a3a33c4&";

// ── XLSX UPLOAD PARSER ───────────────────────────────────────────────────────
async function parseXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array", cellDates:true });
        // Find the sheet with the most rows
        let rows = [];
        for (const name of wb.SheetNames) {
          const j = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval:"", raw:false, dateNF:"dd/mm/yyyy hh:mm:ss" });
          if (j.length > rows.length) rows = j;
        }
        if (!rows.length) { reject(new Error("Spreadsheet appears empty")); return; }

        // Map each row to a raw entry using flexible key matching
        const rawRows = [];
        for (const r of rows) {
          const k = {};
          for (const key of Object.keys(r)) k[key.toLowerCase().trim()] = r[key];

          const ratingVal = k["operation rating"] ?? k["rating"] ?? k["op rating"] ?? "";
          const rating = parseFloat(String(ratingVal));
          if (isNaN(rating) || rating <= 0) continue;

          // Parse timestamp — comes back as string "dd/mm/yyyy hh:mm:ss" or "m/d/yyyy h:mm:ss"
          const tsRaw = String(k["timestamp"] ?? "").trim();
          let date = null;
          if (tsRaw) {
            // Try dd/mm/yyyy hh:mm:ss  (AU Google Forms format)
            const au = tsRaw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            if (au) {
              const [,a,b,y] = au;
              // If a>12 it's definitely day; otherwise ambiguous — AU form so treat as dd/mm
              date = new Date(parseInt(y), parseInt(b)-1, parseInt(a));
            }
            if (!date || isNaN(date)) date = new Date(tsRaw);
          }
          if (!date || isNaN(date)) continue;

          const fbRaw = k["additional information (not required)"] ?? k["additional information"] ?? k["feedback"] ?? k["comments"] ?? "";
          const fb = String(fbRaw).trim();
          const feedback = (fb && fb !== "NULL" && fb.length < 1000) ? [fb] : [];

          rawRows.push({
            date, ms: date.getTime(), rating,
            creator: String(k["op creator"] ?? k["creator"] ?? k["zeus"] ?? "").trim(),
            opName:  String(k["operation name"] ?? k["op name"] ?? k["name"] ?? "").trim(),
            feedback,
          });
        }

        if (!rawRows.length) { reject(new Error("No rows with valid ratings found. Check 'Operation Rating' column.")); return; }
        rawRows.sort((a,b) => a.ms - b.ms);

        // Cluster into sessions (gap > 4h = new op)
        const GAP = 4 * 3600 * 1000;
        const clusters = [];
        for (const row of rawRows) {
          const last = clusters[clusters.length-1];
          if (!last || row.ms - last.lastMs > GAP) clusters.push({ rows:[row], lastMs:row.ms, date:row.date });
          else { last.rows.push(row); last.lastMs = row.ms; }
        }

        // Known op metadata indexed by cluster order (from your pasted data)
        const KNOWN = [
          ["CONTINGENCY","Jag McMuffin"],["CONTINGENCY PT2","Jag McMuffin"],
          ["NOTHERN RISING","Jag McMuffin"],["VALHALLA","Sir Danger"],
          ["AFGHAN OP","Sir Danger"],["VAINS OP","Vain"],
          ["CONTRACT","Jag McMuffin"],["CONTRACT PT2","Jag McMuffin"],
          ["EASTERN UHLAN","Jag McMuffin"],["DUST OFF","Sir Danger"],
          ["Bacons Op","Bacon"],["Black Swan","Jag McMuffin"],
          ["Mountain Snake","Vain"],["Sahara Hunting","Sir Danger"],
          ["Wind Walker","Sir Danger"],["Black Foliage","Jag McMuffin"],
          ["Wait Wait Wait","Jag McMuffin"],["Starkiller","Sir Danger"],
          ["Open Canopy","Jag McMuffin"],["Glass","Sir Danger"],
          ["Night Owl","Jag McMuffin"],["Black Adder","Jag McMuffin"],
          ["Sandstorm","Sir Danger"],["Ocean Master","Sir Danger"],
          ["Golden Serpent","Jag McMuffin"],["Golden Serpent Pt 2","Jag McMuffin"],
          ["Zhnetsy Call","Sir Danger"],["Tribal Welcome","Jag McMuffin"],
          ["El Rey","Jag McMuffin"],["Hermes Gauntlet","Jag McMuffin"],
          ["Stinky Saturdays","Jag McMuffin"],["Condemned Broker","Jag McMuffin"],
          ["Phantom Fury","Sir Danger"],["Milfhunter Fury","Jag McMuffin"],
          ["Iranian Connection","Sir Danger"],["Fjord","Sir Danger"],
          ["RED WATER","Sir Danger"],["EMERALD VEIL I","Jag McMuffin"],
          ["EMERALD VEIL II","Jag McMuffin"],
        ];

        const ops = clusters.map((cl, i) => {
          const sheetName    = cl.rows.find(r => r.opName)?.opName || "";
          const sheetCreator = cl.rows.find(r => r.creator)?.creator || "";
          const [knownName, knownCreator] = KNOWN[i] || [`OP ${i+1}`, "Unknown"];
          const name    = sheetName    || knownName;
          const creator = sheetCreator || knownCreator;
          const ratings = cl.rows.map(r => r.rating);
          const avg = ratings.reduce((a,b)=>a+b,0) / ratings.length;
          const d = cl.date;
          return {
            name, creator,
            year:  String(d.getFullYear()).slice(-2),
            month: d.toLocaleString("default", { month:"long" }),
            rating: parseFloat(avg.toFixed(2)),
            responses: ratings.length,
            feedback: cl.rows.flatMap(r => r.feedback).filter(Boolean),
          };
        });

        resolve(ops);
      } catch(err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("File read failed"));
    reader.readAsArrayBuffer(file);
  });
}

// ── FALLBACK DATA (used when no sheet connected) ─────────────────────────────
// ── REAL DATA from AZO spreadsheet ──────────────────────────────────────────
const FALLBACK_OPS = [
  { name:"CONTINGENCY", creator:"Jag McMuffin", year:"22", month:"March", rating:8.70, responses:10, feedback:[
    "I was lone survivor","Very fun, you guys did a fantastic job, really good polishing for the whole thing. <3",
    "Danger crying as he ran past my body was hilarious","Let people who disconnect reconnect.",
    "Amazing work on the op, loved the teamwork and somewhat comprehendible communication. Amazing atmosphere and the appearance of the demon was perfect, not too prominent, but enough to inflict pure terror.",
    "the enemy and team had the same colour helmet (red) so it was confusing for us but overall it was great",
    "was very chaotic and disorganized, but still fun",
    "Epic map and sounds +10. Maybe limit Zeus intervention healing too much removing role of Medics. Would be good to next time with quicker organisation at start and start at set time.",
    "Pog Champ","The only complaints i have: 1. should have let us get our own guns"
  ]},
  { name:"CONTINGENCY PT2", creator:"Jag McMuffin", year:"22", month:"March", rating:6.12, responses:8, feedback:[
    "Still a good op, great start. Just requires some more polishing off. It would be great to do some more room clearing too.",
    "Would of been better if map wasn't Redux, frames go brr","Frequent frame drops but thats arma, otherwise good op.",
    "good till helicopter crash",
    "Good concept for a continuation of previous op, however i had 2 of 4 squadmates that listened while one wondered off with alpha team",
    "very gay",
    "My perspective I had no clue what was going on. couldnt see anything. Still was a fun thrilling Operation though.",
    "Started off in 8-9 territory, quality of operation slightly reduced later on however amazing start"
  ]},
  { name:"NOTHERN RISING", creator:"Jag McMuffin", year:"22", month:"April", rating:6.43, responses:7, feedback:[
    "the op was mid but wish it could have been a bit more booom bang",
    "mid, medical was dope, people not responding on coms, alright op.",
    "loved it -cute ratboi","Good op, some misbehaviour not zeus fault. Too much zeus intervention",
    "OP was good just some squad members made me need chemotherapy",
    "Also people leaving squads.","I LOVE KAT <3 also rat is great"
  ]},
  { name:"VALHALLA", creator:"Sir Danger", year:"22", month:"April", rating:6.11, responses:9, feedback:[
    "squad two was fucking idiots, kinda mid op","Despite the disgusting comment being made, conduct on my team was good and enjoyed it",
    "badkarma is my dad","teach people how to ace interact, not to get in the pilot seat",
    "Overall, OP was good, liked the idea of it being in the snow (plz no Poland again).",
    "Liked it heeps",
    "needs more conflict, good group plays and splitting off, pretty fun and good use of vehicles",
    "Started off decently well, Squad 1 performed well and everyone listened, op was a bit boring at stages due to lack of objectives"
  ]},
  { name:"AFGHAN OP", creator:"Sir Danger", year:"22", month:"May", rating:6.50, responses:4, feedback:[
    "didnt like how vain kept yelling at everyone for doing nothing",
    "very fun, got to gun cunts with a bushmaster, more scenarios involving EO disposal/diffusal would be better",
    "this op was fun but it was very frustrating when you clear some one then they get shot"
  ]},
  { name:"VAINS OP", creator:"Vain", year:"22", month:"May", rating:6.33, responses:3, feedback:[
    "Liked the op vain but hated everyone there","Really well done man, just needed more people.",
    "Extremely fun op, numbers were an issue but not Vain's fault. Needed people to actually listen to commands rather than go commando."
  ]},
  { name:"CONTRACT", creator:"Jag McMuffin", year:"22", month:"May", rating:8.20, responses:5, feedback:[
    "i got to shoo Arabians","to much fog very slow, lost",
    "Too many sandstorms, 1 or 2 is enough, not for literally every scenario/engagement during an op",
    "a lot of explosions, and could not see anything"
  ]},
  { name:"CONTRACT PT2", creator:"Jag McMuffin", year:"22", month:"June", rating:8.67, responses:6, feedback:[
    "cool","fuun",
    "liked the teamwork. vain needs a leash. great fun",
    "Good usage of ambience (rain), train Ronan before testing him, Schnitzel and Satan should be next in line for CPL",
    "the combat and overall operation was very fun and engaging. lots of fun"
  ]},
  { name:"EASTERN UHLAN", creator:"Jag McMuffin", year:"22", month:"June", rating:9.50, responses:4, feedback:[
    "unintentionally most immersive op","Loved chanting wendigo praises in latin <3 -jager",
    "I pooped on more than 1 occasion",
    "I shat my fucking pants, amazing job improvising with the OP and just overall amazing tension building."
  ]},
  { name:"DUST OFF", creator:"Sir Danger", year:"22", month:"September", rating:8.57, responses:7, feedback:[
    "bit chaotic, got left behind a few times. Otherwise, fun first Op :D",
    "Confusing pacing but overall atmosphere had great start and ending, advice might need to give us like an objective/reason why we were fighting.",
    "very good op, would of liked to see more people and some sort of air unit.",
    "The Dust is always a little too much. lack of communication between orders. good fun. A nice amount of enemies. A fast and efficient mission"
  ]},
  { name:"Bacons Op", creator:"Bacon", year:"22", month:"September", rating:5.50, responses:2, feedback:[
    "boring, too much combat all at once, not much of a briefing, objectives close together"
  ]},
  { name:"Black Swan", creator:"Jag McMuffin", year:"22", month:"November", rating:8.20, responses:5, feedback:[
    "a bit chaotic on command structure, but there were very few of us. more sandbox approach is good",
    "Enjoyed it",
    "Enjoyed the op, just a bit of a failure of the chain of command towards the end",
    "I personally really liked the op well done, I think that running in pairs was really fun.",
    "Good vibe, definitely more OPs needed like this in terms of freedom and realism but behavior needs to be better and less night time and more environment variety"
  ]},
  { name:"Mountain Snake", creator:"Vain", year:"22", month:"November", rating:7.25, responses:4, feedback:[
    "Was shit cuz I die of friendly fire Michael Jordan nade",
    "felt short, 1 objective, didnt know it was 1 life, good fight, town fights good"
  ]},
  { name:"Sahara Hunting", creator:"Sir Danger", year:"22", month:"December", rating:7.75, responses:4, feedback:[
    "A good mess, I like having more general objectives we can charge into for fun",
    "I really love the balance between milsim and fucking around. This coupled with the well thought out plans and lore creates a fun experience even if shit goes wrong here and there.",
    "Was a fun op I liked my squad with bacon and the freedom was very good"
  ]},
  { name:"Wind Walker", creator:"Sir Danger", year:"22", month:"December", rating:9.00, responses:5, feedback:[
    "I kept dying.",
    "Enjoyed playing recon and setting up ranges and stuff, Would like a bit more of a command structure for different squads and command groups maybe to keep us all together",
    "Really nice and in depth operation. Recon was pretty boring but it was vital so I didn't mind it too much. The night time raid on that red zone was pretty sick.",
    "A good op tonight, Teamwork was amazing not too much chatter and no bullshit happened tonight was gud<3"
  ]},
  { name:"Black Foliage", creator:"Jag McMuffin", year:"23", month:"February", rating:7.75, responses:4, feedback:[
    "So fun I loved it",
    "Liked the setup and map, good mission :D",
    "I don't particularly like the map of Preacher but it was pretty fun, less night missions more daytime.",
    "Overall it was a good op and I really liked the part where you took control of the enemy's and made us negotiate with them and it was a really cool idea. I also loved how lambs AI made us able to sneak around.",
    "Lots of sexy kit in the arsenal. Nice map, very atmospheric. Was fun sneaking around on that with NODs off."
  ]},
  { name:"Wait Wait Wait", creator:"Jag McMuffin", year:"23", month:"February", rating:7.00, responses:4, feedback:[
    "NO NO NO NO NO WAIT WAIT WAIT WAIT","You fucking cunt this was just a meme operation",
    "NO NO NO NO NO. WAIT WAIT WAIT WAIT WAIT."
  ]},
  { name:"Starkiller", creator:"Sir Danger", year:"23", month:"March", rating:6.75, responses:4, feedback:[
    "A bit out of sync",
    "Was a very good op first bit was decently cinematic. Pls make AI less tanky<3",
    "Was alright, not too exciting or cool, not overly boring, quite short?",
    "Loved the setting and ideas, AI were a bit terminatory",
    "Love the fresh approach to a more milsim and tactical experience rather than the usual fooling around."
  ]},
  { name:"Open Canopy", creator:"Jag McMuffin", year:"23", month:"March", rating:7.67, responses:3, feedback:[
    "Was a good op, enjoyed having a smaller team was very coordinated",
    "More on the serious side. Actually quite enjoyable and the amount of hostiles were very reasonable so fighting them off wasn't too bad yet still very scary.",
    "The hot extract was fun."
  ]},
  { name:"Glass", creator:"Sir Danger", year:"23", month:"April", rating:7.50, responses:4, feedback:[
    "Op was good :) Too many AI at the start but we chillin",
    "Thermal next time, prefer desert op, next time we play as OPFOR?",
    "Liked the setting and stuffs, team was a bit eh, Maybe we should have some proper team training",
    "Overall good OP, well planned out and executed, shame there are too many retards in the unit."
  ]},
  { name:"Night Owl", creator:"Jag McMuffin", year:"23", month:"April", rating:9.00, responses:4, feedback:[
    "Funny op cheesy",
    "REALLY good op, Civil was the real MVP rip<3, Rlly cinematic and lore was made today",
    "The hostiles were like overpowering but not overpowering at the same time. Like halfway through the OP, the part with the enemy air support dunking on us was very exciting and scary.",
    "The operation was actually fun for me, I was able to communicate and cooperate with my team. Good teamwork, coordination and overall very fun for my experience"
  ]},
  { name:"Black Adder", creator:"Jag McMuffin", year:"23", month:"April", rating:7.75, responses:4, feedback:[
    "i shot a guy with an rpg and he lived lol",
    "Was a rather boring op half way through but the start was fun with the stealth and my mouse wheel fucking broke",
    "Sick operation but lotta walking. My legs are hurting irl.",
    "Quite enjoyable, really enjoying the enemies going down after one shot. Could be a little more Cinematic but that's just a unimportant nit pick."
  ]},
  { name:"Sandstorm", creator:"Sir Danger", year:"23", month:"May", rating:8.50, responses:6, feedback:[
    "Op was pretty cool and adrenaline filling with the phonk. too much AI at the start. CLOSE air support was cool",
    "Good op and shit but I heavily under estimated this whole 1 life thing.",
    "i came in my pants playing dis op it was very fun",
    "danger sex so good !!!"
  ]},
  { name:"Ocean Master", creator:"Sir Danger", year:"23", month: "June", rating:5.00, responses:2, feedback:[
    "is was alright I guess",
    "Is what is bruh that defend part was funny and like I had lung cancer Walter whit"
  ]},
  { name:"Golden Serpent", creator:"Jag McMuffin", year:"23", month:"October", rating:8.14, responses:3, feedback:[
    "i was only there for like 40 mins",
    "Enjoyed for a good start back, urban combat ambushes are hellish, map maybe a bit buggy at points, good amount of hostiles and difficulty.",
    "Straight to the point OBJ, sick urban warfare and the props and cars on the street made the place feel very dense and real. Presence of civilians added a nice touch as we actually had to pay attention before shooting."
  ]},
  { name:"Golden Serpent Pt 2", creator:"Jag McMuffin", year:"23", month:"October", rating:8.14, responses:7, feedback:[
    "Enemy sniper too op (he the best)",
    "Have a clearer objective",
    "Op was perfect, maybe be less hands on w/ the heals or train the medics to the point where it's not needed. Loved the action.",
    "Would like less laggy map, enjoyed the high tempo urban warfare, enjoyed it felt very Blackhawk down ish.",
    "was good op :), made by jagy poo",
    "i like shoot","Enemy sniper best (everyone else have skill issues)"
  ]},
  { name:"Zhnetsy Call", creator:"Sir Danger", year:"23", month:"October", rating:7.86, responses:7, feedback:[
    "Great op, only problems I had was communication between squads like the squad that kept running off and ramboing.",
    "i loved the Op i wanna do more like this :)",
    "Could have used a more thought out plan rather than Napoleonic style frontal-attack.",
    "no clear chain of command therefore confusion. lack of medical caused an arse of a problem. lack of coordination led to casualties",
    "My only real concern is how RPL's blue team seemed to be too disconnected from all the action.",
    "Operation itself was pretty good but lack of plans/information and communication kind of led us to shit a little. Frames weren't amazing and the terrain had a few problems."
  ]},
  { name:"Tribal Welcome", creator:"Jag McMuffin", year:"23", month:"November", rating:5.00, responses:4, feedback:[
    "The op was really mid and the configs being broke made it worse. 7B protocol was really fun though",
    "Server settings need to be set properly, op was a bit confusing story wise, too many anomalies in some places which just caused it to be just a minefield of meatgrinders.",
    "9mm seemed a bit weak especially if we're doing PVP at the end.",
    "Was too scuffed Very low FPS 2 people quit because of fps wouldnt do again with those ace settings"
  ]},
  { name:"El Rey", creator:"Jag McMuffin", year:"23", month:"November", rating:7.67, responses:6, feedback:[
    "Very good","jag was hot :), was a good op(jag need to make more ops)",
    "Very good, felt really atmospheric at times however lag needs to be improved, I would recommend more scripts to increase the cinematic feel.",
    "Perfect amount of hostiles I reckon. Villages didn't feel packed nor empty and the amount of hostiles towards the end with the evac and all that wasn't too much to the point where it was near impossible but enough for us to actually have a challenge."
  ]},
  { name:"Hermes Gauntlet", creator:"Jag McMuffin", year:"23", month:"November", rating:7.00, responses:2, feedback:[
    "Very challenging but really sick ambushes especially the MG posted on the dam shooting down on us out of nowhere. So good it was suppressing me irl.",
    "Very laggy, maybe more customization with ops that are less serious."
  ]},
  { name:"Stinky Saturdays", creator:"Jag McMuffin", year:"24", month:"February", rating:7.50, responses:2, feedback:[
    "Very intense and challenging gun fights! Maybe a tad bit overwhelming but still super fun!!",
    "Op was good, communication was poor, tactics were questionable to say the least and there felt like there was no command. Op concept was great and interesting, execution by players was terrible."
  ]},
  { name:"Condemned Broker", creator:"Jag McMuffin", year:"24", month:"May", rating:10.00, responses:2, feedback:[
    "Good op","Excellent op love, had a very fun time and was very cinematic towards the end <3"
  ]},
  { name:"Phantom Fury", creator:"Sir Danger", year:"24", month:"October", rating:8.50, responses:6, feedback:[
    "very good","AZO IS SO FUCKINBG BACKKKKJK","Day > Night","IMPASSIBAL!!!!!",
    "i'm fucking every trans person i see"
  ]},
  { name:"Milfhunter Fury", creator:"Jag McMuffin", year:"24", month:"November", rating:7.50, responses:2, feedback:[
    "Im very tired","very good, had a great time, love pookie danger"
  ]},
  { name:"Iranian Connection", creator:"Sir Danger", year:"24", month:"November", rating:8.56, responses:9, feedback:[
    "bing chilling very good","Good op, very fun and thrilling",
    "bro I joined at the end and it was still fun","come to join op tmr, love you jag, mwa",
    "Im horny rn - malt",
    "i was more scared on my exam then this op, i also nearly pissed myself this afternoon. 10/10 very good op",
    "very funny, next time maybe in a more enclosed map",
    "Was good how you let the groups operate separately despite that not being intended. was my first op but was a tonne of fun"
  ]},
  { name:"Fjord", creator:"Sir Danger", year:"24", month:"November", rating:9.40, responses:5, feedback:[
    "Mr stranger danger didn't give me the smooch :(",
    "Very good jag is my dad he drop me off at fire station pls let me jump out of airplane bendy is gay",
    "w op, the cia stuff was cool, the video was kinda cringe but, every thing els was great",
    "excluding the bug of falling though the ship, fantastic op really enjoyed it. awesome work.",
    "I give you good rating Mr Eamonn, only downrated for falling through ship issues, all objectives showing on map, and just some confusion between command and ground."
  ]},
  { name:"RED WATER", creator:"Sir Danger", year:"24", month:"December", rating:7.75, responses:4, feedback:[
    "no danger tonight but that's ok because jag is a mega sigma chat man",
    "Jag died very sad :(",
    "VERY EPIC we will die in the forest",
    "Simply cos of arma's ai it was a little funny, but the concept and the overall execution was pretty well done and enjoyable, just arma being arma cant do nothing about it"
  ]},
  { name:"EMERALD VEIL I", creator:"Sir Danger", year:"25", month:"March", rating:7.00, responses:2, feedback:[
    "Was fun needed more ppl and a full brief w/ pre set cmdr"
  ]},
  { name:"EMERALD VEIL II", creator:"Jag McMuffin", year:"26", month:"May", rating:7.17, responses:12, feedback:[
    "Confusion, Too many Dead fr. But pretty good ngl.",
    "Enjoyed the Cinematic opening, loadouts felt good maybe a little more ammo or even airdrops would be nice. Ending was a little flag the regroup was too far and should of ended after the firefight.",
    "smh helicopter hit tree, also squad cohesion sucked. Casualty management and general communication with and between squads sucked",
    "Great concept and mission environment for the op, loved the realistic op brief document, however, the rough insertion for Talon and the hardly functional Teamspeak situation about halfway through the op tanked my enjoyment.",
    "great to be back, massive delay at start couldve been prevented by doing earlier start. otherwise great",
    "ACRE failing was the biggest issue. Fog at airbase made NVGs unusable but AI seemed to have good visibility. Use of AI smokes seemed excessive.",
    "NEVER make me a squad leader again. Op wouldve been okay if I didnt have a retarded team and all the buggy tfar shit didnt help, enemies too tanky"
  ]},
];

const CREATORS = {
  "Jag McMuffin": { color: "#3B82F6", bg: "#1e3a5f", initials: "JM" },
  "Sir Danger":   { color: "#F97316", bg: "#5c2c00", initials: "SD" },
  "Vain":         { color: "#A855F7", bg: "#3b1f5e", initials: "VN" },
  "Bacon":        { color: "#22C55E", bg: "#1a3d1a", initials: "BC" },
};

function ratingColor(r) {
  if (r >= 9) return "#22C55E";
  if (r >= 7.5) return "#84CC16";
  if (r >= 6) return "#F59E0B";
  return "#EF4444";
}
function RatingBar({ value }) {
  return <div style={{ position:"relative", height:6, background:"#1e2d42", borderRadius:3, overflow:"hidden", width:"100%" }}><div style={{ position:"absolute", left:0, top:0, height:"100%", width:`${(value/10)*100}%`, background:ratingColor(value), borderRadius:3 }} /></div>;
}
function RatingBadge({ value }) {
  return <span style={{ fontFamily:"monospace", fontSize:13, fontWeight:700, color:ratingColor(value), background:"rgba(0,0,0,0.4)", padding:"2px 8px", borderRadius:4, border:`1px solid ${ratingColor(value)}44` }}>{value.toFixed(2)}</span>;
}
function creatorStats(name, ops) {
  const o = ops.filter(x => x.creator === name);
  const avg = o.reduce((s,x)=>s+x.rating,0)/o.length;
  return { ops:o.length, avg:avg.toFixed(2), best:[...o].sort((a,b)=>b.rating-a.rating)[0], worst:[...o].sort((a,b)=>a.rating-b.rating)[0], totalResponses:o.reduce((s,x)=>s+x.responses,0), allOps:o };
}

// SheetJS from CDN — loaded via script tag in App

// ── DASHBOARD ────────────────────────────────────────────────────────────────
function DashboardView({ ops }) {
  const totalResponses = ops.reduce((s,o)=>s+o.responses,0);
  const overallAvg = (ops.reduce((s,o)=>s+o.rating,0)/ops.length).toFixed(2);
  const topOp = [...ops].sort((a,b)=>b.rating-a.rating)[0];
  const years = [...new Set(ops.map(o=>o.year))].sort();
  const yearData = years.map(yr => {
    const yo = ops.filter(o=>o.year===yr);
    return { year:`20${yr}`, avg:parseFloat((yo.reduce((s,o)=>s+o.rating,0)/yo.length).toFixed(2)), ops:yo.length };
  });
  const recentOps = [...ops].slice(-6).reverse();
  // Activity leaderboard — ops ranked by response volume
  const activityRanked = [...ops].sort((a,b) => b.responses - a.responses);
  const maxResponses = activityRanked[0]?.responses || 1;

  return (
    <div>
      {/* Activity leaderboard */}
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:4 }}>🏟 ENGAGEMENT LEADERBOARD</div>
        <div style={{ fontFamily:"sans-serif", fontSize:11, color:"#3d5a7a", marginBottom:14 }}>Ops ranked by volume of feedback responses</div>
        {activityRanked.map((op, i) => {
          const cs = CREATORS[op.creator];
          const pct = (op.responses / maxResponses) * 100;
          const medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : `#${i+1}`;
          return (
            <div key={op.name+op.year} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
              <div style={{ width:28, fontFamily:"monospace", fontSize:i<3?14:11, textAlign:"center", flexShrink:0 }}>{medal}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3, alignItems:"center" }}>
                  <span style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:"#c8d8e8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{op.name}</span>
                  <div style={{ display:"flex", gap:8, flexShrink:0, marginLeft:8, alignItems:"center" }}>
                    <span style={{ fontFamily:"monospace", fontSize:11, color:ratingColor(op.rating) }}>{op.rating.toFixed(2)}</span>
                    <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700, color:"#F59E0B" }}>{op.responses}</span>
                  </div>
                </div>
                <div style={{ position:"relative", height:5, background:"#1e2d42", borderRadius:3, overflow:"hidden" }}>
                  <div style={{ position:"absolute", left:0, top:0, height:"100%", width:`${pct}%`, background: i===0?"#F59E0B": i===1?"#9CA3AF": i===2?"#B45309": cs?.color||"#3B82F6", borderRadius:3 }}/>
                </div>
                <div style={{ fontFamily:"sans-serif", fontSize:10, color:cs?.color||"#5b7fa6", marginTop:2 }}>{op.creator} · 20{op.year}</div>
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(150px, 1fr))", gap:12, marginBottom:24 }}>
        {[{label:"TOTAL OPS",value:ops.length,sub:"all time"},{label:"OVERALL AVG",value:overallAvg,sub:"out of 10.00",color:ratingColor(parseFloat(overallAvg))},{label:"RESPONSES",value:totalResponses,sub:"total feedback"},{label:"TOP RATED",value:topOp?.rating.toFixed(2),sub:topOp?.name,color:"#22C55E"}].map(({label,value,sub,color})=>(
          <div key={label} style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:"16px 18px" }}>
            <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.12em", color:"#5b7fa6", marginBottom:4 }}>{label}</div>
            <div style={{ fontFamily:"monospace", fontSize:26, fontWeight:700, color:color||"#e8edf2", lineHeight:1.1 }}>{value}</div>
            <div style={{ fontFamily:"sans-serif", fontSize:11, color:"#4a6581", marginTop:4, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{sub}</div>
          </div>
        ))}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:24 }}>
        <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16 }}>
          <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>AVG RATING BY YEAR</div>
          <ResponsiveContainer width="100%" height={150}><LineChart data={yearData}><CartesianGrid strokeDasharray="3 3" stroke="#1a2d42"/><XAxis dataKey="year" tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false}/><YAxis domain={[5,10]} tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false} width={24}/><Tooltip contentStyle={{background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:6,fontFamily:"monospace"}} itemStyle={{color:"#e8edf2"}} labelStyle={{color:"#5b7fa6"}}/><Line type="monotone" dataKey="avg" stroke="#3B82F6" strokeWidth={2} dot={{fill:"#3B82F6",r:3}}/></LineChart></ResponsiveContainer>
        </div>
        <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16 }}>
          <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>OPS PER YEAR</div>
          <ResponsiveContainer width="100%" height={150}><BarChart data={yearData}><CartesianGrid strokeDasharray="3 3" stroke="#1a2d42" vertical={false}/><XAxis dataKey="year" tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false} width={20}/><Tooltip contentStyle={{background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:6,fontFamily:"monospace"}}/><Bar dataKey="ops" radius={[4,4,0,0]}>{yearData.map((_,i)=><Cell key={i} fill={["#3B82F6","#F97316","#A855F7","#22C55E","#F59E0B","#EF4444"][i%6]}/>)}</Bar></BarChart></ResponsiveContainer>
        </div>
      </div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:24 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:14 }}>ZEUS COMMANDER LEADERBOARD</div>
        {Object.entries(CREATORS).filter(([n])=>ops.some(o=>o.creator===n)).map(([name,style],i)=>{
          const stats = creatorStats(name, ops);
          return (
            <div key={name} style={{ display:"flex", alignItems:"center", gap:12, marginBottom:i<Object.keys(CREATORS).length-1?14:0 }}>
              <div style={{ width:20, fontFamily:"monospace", fontSize:11, color:"#3d5a7a", textAlign:"center" }}>#{i+1}</div>
              <div style={{ width:34, height:34, borderRadius:"50%", background:style.bg, border:`2px solid ${style.color}`, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700, fontSize:11, color:style.color, flexShrink:0 }}>{style.initials}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                  <span style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:13, color:"#c8d8e8" }}>{name}</span>
                  <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700, color:ratingColor(parseFloat(stats.avg)) }}>{stats.avg}</span>
                </div>
                <RatingBar value={parseFloat(stats.avg)}/>
                <div style={{ display:"flex", gap:10, marginTop:3 }}>
                  <span style={{ fontFamily:"sans-serif", fontSize:10, color:"#3d5a7a" }}>{stats.ops} ops</span>
                  <span style={{ fontFamily:"sans-serif", fontSize:10, color:"#3d5a7a" }}>{stats.totalResponses} responses</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:14 }}>RECENT OPERATIONS</div>
        {recentOps.map(op=>(
          <div key={op.name} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"8px 0", borderBottom:"1px solid #0e1f30" }}>
            <div>
              <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:13, color:"#c8d8e8" }}>{op.name}</div>
              <div style={{ fontFamily:"sans-serif", fontSize:11, color:CREATORS[op.creator]?.color||"#5b7fa6" }}>{op.creator} · 20{op.year}</div>
            </div>
            <RatingBadge value={op.rating}/>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── OPS ──────────────────────────────────────────────────────────────────────
function OperationsView({ ops, onSelectOp }) {
  const [filter, setFilter] = useState("ALL");
  const [sortBy, setSortBy] = useState("year");
  const [search, setSearch] = useState("");
  const years = ["ALL", ...[...new Set(ops.map(o=>o.year))].sort()];
  let filtered = filter==="ALL" ? ops : ops.filter(o=>o.year===filter);
  if (search) filtered = filtered.filter(o=>o.name.toLowerCase().includes(search.toLowerCase())||o.creator.toLowerCase().includes(search.toLowerCase()));
  if (sortBy==="rating") filtered=[...filtered].sort((a,b)=>b.rating-a.rating);
  else if (sortBy==="year") filtered=[...filtered].reverse();
  else if (sortBy==="responses") filtered=[...filtered].sort((a,b)=>b.responses-a.responses);
  return (
    <div>
      <div style={{ display:"flex", gap:8, marginBottom:12, flexWrap:"wrap" }}>
        <input placeholder="Search ops or creators…" value={search} onChange={e=>setSearch(e.target.value)} style={{ flex:1, minWidth:130, background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:6, padding:"8px 12px", color:"#c8d8e8", fontFamily:"sans-serif", fontSize:12, outline:"none" }}/>
        <select value={sortBy} onChange={e=>setSortBy(e.target.value)} style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:6, padding:"8px 10px", color:"#c8d8e8", fontFamily:"sans-serif", fontSize:12, outline:"none" }}>
          <option value="year">Latest First</option>
          <option value="rating">Top Rated</option>
          <option value="responses">Most Feedback</option>
        </select>
      </div>
      <div style={{ display:"flex", gap:6, marginBottom:16, flexWrap:"wrap" }}>
        {years.map(yr=>(
          <button key={yr} onClick={()=>setFilter(yr)} style={{ padding:"5px 12px", borderRadius:5, border:filter===yr?"1px solid #3B82F6":"1px solid #1e3a5f", background:filter===yr?"#1e3a5f":"transparent", color:filter===yr?"#3B82F6":"#5b7fa6", fontFamily:"monospace", fontWeight:700, fontSize:11, cursor:"pointer" }}>
            {yr==="ALL"?"ALL":`20${yr}`}
          </button>
        ))}
      </div>
      <div style={{ fontSize:10, fontFamily:"monospace", color:"#3d5a7a", marginBottom:12, letterSpacing:"0.08em" }}>{filtered.length} OPERATIONS</div>
      {filtered.map(op=>{
        const cs = CREATORS[op.creator];
        return (
          <div key={op.name+op.year} style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:"12px 14px", marginBottom:8, cursor:"pointer" }} onClick={()=>onSelectOp(op)}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:7 }}>
              <div>
                <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:14, color:"#e8edf2" }}>{op.name}</div>
                <div style={{ display:"flex", gap:8, alignItems:"center", marginTop:2 }}>
                  <span style={{ fontFamily:"sans-serif", fontSize:11, color:cs?.color||"#5b7fa6", fontWeight:700 }}>{op.creator}</span>
                  <span style={{ fontFamily:"sans-serif", fontSize:10, color:"#3d5a7a" }}>· 20{op.year} · {op.month}</span>
                  <span style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a" }}>{op.responses}↩</span>
                </div>
              </div>
              <RatingBadge value={op.rating}/>
            </div>
            <RatingBar value={op.rating}/>
            {op.feedback?.length>0 && <div style={{ fontFamily:"sans-serif", fontSize:10, color:"#3d5a7a", marginTop:6 }}>"{op.feedback[0]}"</div>}
          </div>
        );
      })}
    </div>
  );
}

// ── FEEDBACK TAB ─────────────────────────────────────────────────────────────
function FeedbackView({ ops, selectedOp: initOp }) {
  const [selected, setSelected] = useState(initOp || ops[ops.length-1]);
  const [search, setSearch] = useState("");
  const opsSorted = [...ops].reverse();
  const hasFeedback = ops.filter(o=>o.feedback?.length>0);

  return (
    <div style={{ display:"grid", gridTemplateColumns:"220px 1fr", gap:12, minHeight:400 }}>
      {/* Op list */}
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, overflow:"hidden", display:"flex", flexDirection:"column", maxHeight:"80vh" }}>
        <div style={{ padding:"10px 12px", borderBottom:"1px solid #1e3a5f" }}>
          <input placeholder="Filter ops…" value={search} onChange={e=>setSearch(e.target.value)} style={{ width:"100%", background:"#060f1e", border:"1px solid #1e3a5f", borderRadius:5, padding:"6px 8px", color:"#c8d8e8", fontFamily:"sans-serif", fontSize:11, outline:"none", boxSizing:"border-box" }}/>
        </div>
        <div style={{ overflowY:"auto", flex:1 }}>
          {opsSorted.filter(o=>o.feedback?.length>0&&(!search||o.name.toLowerCase().includes(search.toLowerCase()))).map(op=>(
            <div key={op.name+op.year} onClick={()=>setSelected(op)} style={{ padding:"9px 12px", borderBottom:"1px solid #0e1f30", cursor:"pointer", background:selected?.name===op.name&&selected?.year===op.year?"#1e3a5f":"transparent" }}>
              <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:"#c8d8e8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{op.name}</div>
              <div style={{ display:"flex", justifyContent:"space-between", marginTop:2 }}>
                <span style={{ fontFamily:"sans-serif", fontSize:10, color:CREATORS[op.creator]?.color||"#5b7fa6" }}>{op.creator}</span>
                <span style={{ fontFamily:"monospace", fontSize:10, color:ratingColor(op.rating) }}>{op.rating.toFixed(2)}</span>
              </div>
            </div>
          ))}
          {hasFeedback.length===0&&<div style={{ padding:16, fontFamily:"sans-serif", fontSize:11, color:"#3d5a7a", textAlign:"center" }}>No feedback data yet.<br/>Connect a Google Sheet or add feedback to ops.</div>}
        </div>
      </div>
      {/* Feedback content */}
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, overflowY:"auto", maxHeight:"80vh" }}>
        {selected ? (
          <>
            <div style={{ marginBottom:14 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                <div>
                  <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:16, color:"#e8edf2" }}>{selected.name}</div>
                  <div style={{ fontFamily:"sans-serif", fontSize:12, color:CREATORS[selected.creator]?.color||"#5b7fa6", marginTop:2 }}>{selected.creator} · 20{selected.year} {selected.month&&`· ${selected.month}`}</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <RatingBadge value={selected.rating}/>
                  <div style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a", marginTop:4 }}>{selected.responses} responses</div>
                </div>
              </div>
              <div style={{ marginTop:10 }}><RatingBar value={selected.rating}/></div>
            </div>
            <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:10 }}>FEEDBACK RESPONSES ({selected.feedback?.length||0})</div>
            {selected.feedback?.length>0 ? selected.feedback.map((f,i)=>(
              <div key={i} style={{ background:"#060f1e", border:"1px solid #1e3a5f", borderRadius:6, padding:"10px 12px", marginBottom:8 }}>
                <div style={{ display:"flex", gap:8, alignItems:"flex-start" }}>
                  <span style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a", paddingTop:2, flexShrink:0 }}>#{i+1}</span>
                  <span style={{ fontFamily:"sans-serif", fontSize:13, color:"#c8d8e8", lineHeight:1.5 }}>"{f}"</span>
                </div>
              </div>
            )) : <div style={{ fontFamily:"sans-serif", fontSize:12, color:"#3d5a7a", textAlign:"center", padding:24 }}>No feedback recorded for this operation.</div>}
          </>
        ) : <div style={{ fontFamily:"sans-serif", fontSize:12, color:"#3d5a7a", textAlign:"center", padding:40 }}>Select an op to view feedback</div>}
      </div>
    </div>
  );
}

// ── CREATORS VIEW ─────────────────────────────────────────────────────────────
function CreatorsView({ ops }) {
  const activeCreators = Object.entries(CREATORS).filter(([n])=>ops.some(o=>o.creator===n));
  const [selected, setSelected] = useState(activeCreators[0]?.[0]||"");
  const stats = selected ? creatorStats(selected, ops) : null;
  const cs = CREATORS[selected];
  const chartData = stats?.allOps.map(o=>({ name:o.name.length>12?o.name.slice(0,12)+"…":o.name, rating:o.rating, fullName:o.name })) || [];
  return (
    <div>
      <div style={{ display:"flex", gap:8, marginBottom:20, flexWrap:"wrap" }}>
        {activeCreators.map(([name,style])=>{
          const st = creatorStats(name, ops);
          return (
            <button key={name} onClick={()=>setSelected(name)} style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 14px", borderRadius:8, border:selected===name?`1px solid ${style.color}`:"1px solid #1e3a5f", background:selected===name?style.bg:"transparent", cursor:"pointer" }}>
              <div style={{ width:28, height:28, borderRadius:"50%", background:style.bg, border:`2px solid ${style.color}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:700, color:style.color }}>{style.initials}</div>
              <div>
                <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:selected===name?style.color:"#c8d8e8" }}>{name}</div>
                <div style={{ fontFamily:"monospace", fontSize:10, color:ratingColor(parseFloat(st.avg)) }}>{st.avg}</div>
              </div>
            </button>
          );
        })}
      </div>
      {stats && cs && (
        <>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10, marginBottom:20 }}>
            {[{label:"OPS RUN",value:stats.ops},{label:"AVG RATING",value:stats.avg,color:ratingColor(parseFloat(stats.avg))},{label:"RESPONSES",value:stats.totalResponses}].map(({label,value,color})=>(
              <div key={label} style={{ background:"#0d1e33", border:`1px solid ${cs.color}22`, borderRadius:8, padding:"12px 14px", textAlign:"center" }}>
                <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:4 }}>{label}</div>
                <div style={{ fontFamily:"monospace", fontSize:22, fontWeight:700, color:color||cs.color }}>{value}</div>
              </div>
            ))}
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:20 }}>
            {[{label:"🏆 BEST OP",op:stats.best},{label:"📉 LOWEST OP",op:stats.worst}].map(({label,op})=>(
              <div key={label} style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:"12px 14px" }}>
                <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:6 }}>{label}</div>
                <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:"#c8d8e8", marginBottom:6 }}>{op?.name}</div>
                <RatingBadge value={op?.rating||0}/>
              </div>
            ))}
          </div>
          <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
            <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>OP RATINGS HISTORY</div>
            <ResponsiveContainer width="100%" height={170}><BarChart data={chartData} barSize={13}><CartesianGrid strokeDasharray="3 3" stroke="#1a2d42" vertical={false}/><XAxis dataKey="name" tick={{fill:"#5b7fa6",fontSize:8,fontFamily:"monospace"}} axisLine={false} tickLine={false} interval={0} angle={-40} textAnchor="end" height={48}/><YAxis domain={[0,10]} tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false} width={22}/><Tooltip contentStyle={{background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:6,fontFamily:"monospace"}} formatter={(v,n,p)=>[v.toFixed(2),p.payload.fullName]} labelFormatter={()=>""}/><Bar dataKey="rating" radius={[3,3,0,0]}>{chartData.map((d,i)=><Cell key={i} fill={ratingColor(d.rating)}/>)}</Bar></BarChart></ResponsiveContainer>
          </div>
          {stats.allOps.map(op=>(
            <div key={op.name} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"9px 0", borderBottom:"1px solid #0e1f30" }}>
              <div>
                <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:13, color:"#c8d8e8" }}>{op.name}</div>
                <div style={{ fontFamily:"sans-serif", fontSize:10, color:"#3d5a7a" }}>20{op.year} · {op.month} · {op.responses} responses</div>
              </div>
              <RatingBadge value={op.rating}/>
            </div>
          ))}
        </>
      )}
    </div>
  );
}

// ── AI INTEL VIEW ─────────────────────────────────────────────────────────────
const CUTOFF_MONTHS = 12;

function getOpAge(op) {
  const yr = parseInt("20"+op.year);
  const monthMap = {january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
  const mo = monthMap[(op.month||"").toLowerCase()] ?? 6;
  return new Date(yr, mo);
}

function isRecent(op) {
  const now = new Date(2026, 4, 10); // May 2026
  const opDate = getOpAge(op);
  const diffMonths = (now.getFullYear()-opDate.getFullYear())*12 + (now.getMonth()-opDate.getMonth());
  return diffMonths <= CUTOFF_MONTHS;
}

function IntelView({ ops }) {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");
  const [selectedCreator, setSelectedCreator] = useState("ALL");

  const recentOps = ops.filter(isRecent);
  const recentFeedback = recentOps.flatMap(o=>(o.feedback||[]).map(f=>({ text:f, op:o.name, creator:o.creator, rating:o.rating })));
  const allFeedback = ops.flatMap(o=>(o.feedback||[]).map(f=>({ text:f, op:o.name, creator:o.creator, rating:o.rating })));

  const activeCreators = ["ALL", ...Object.keys(CREATORS).filter(n=>ops.some(o=>o.creator===n))];

  async function runAnalysis() {
    setLoading(true); setErr(""); setAnalysis(null);
    const filtered = selectedCreator==="ALL" ? allFeedback : allFeedback.filter(f=>f.creator===selectedCreator);
    const recentFiltered = selectedCreator==="ALL" ? recentFeedback : recentFeedback.filter(f=>f.creator===selectedCreator);

    const prompt = `You are an analyst for AZO (Australian Zeus Ops), an Arma 3 milsim group. Analyse the following op feedback and return ONLY valid JSON.

ALL FEEDBACK (${filtered.length} items):
${filtered.map(f=>`[${f.op} | ${f.creator} | ${f.rating}/10]: "${f.text}"`).join("\n")}

RECENT FEEDBACK (last 12 months, ${recentFiltered.length} items):
${recentFiltered.map(f=>`[${f.op} | ${f.creator} | ${f.rating}/10]: "${f.text}"`).join("\n")}

Return ONLY this JSON structure, no markdown, no explanation:
{
  "strengths": [
    { "title": "string", "description": "string (2 sentences)", "status": "consistent", "evidence": "string (specific example from feedback)", "opCount": number }
  ],
  "weaknesses": [
    { "title": "string", "description": "string (2 sentences)", "status": "active|improving|resolved", "evidence": "string (specific example from feedback)", "opCount": number, "lastSeen": "string (op name)" }
  ],
  "summary": "string (3-4 sentences overall assessment)",
  "trend": "improving|stable|declining",
  "highlights": ["string","string","string"]
}

Keep descriptions to 1 sentence max. Keep evidence quotes under 10 words. Identify exactly 5 strengths and 5 weaknesses. Be concise.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:1000, messages:[{role:"user",content:prompt}] })
      });
      const data = await res.json();
      const text = data.content?.filter(c=>c.type==="text").map(c=>c.text).join("") || "";
      let clean = text.replace(/```json|```/g,"").trim();
      // If JSON got truncated, attempt to close it safely
      if (!clean.endsWith("}")) {
        // Strip trailing incomplete entry and close arrays/object
        clean = clean.replace(/,\s*$/, "");
        clean = clean.replace(/,?\s*\{[^}]*$/, ""); // remove last incomplete object
        const opens = (clean.match(/\[/g)||[]).length - (clean.match(/\]/g)||[]).length;
        const braces = (clean.match(/\{/g)||[]).length - (clean.match(/\}/g)||[]).length;
        clean += "]".repeat(Math.max(0,opens)) + "}".repeat(Math.max(0,braces));
      }
      const parsed = JSON.parse(clean);
      setAnalysis(parsed);
    } catch(e) { setErr("Analysis failed: "+e.message); }
    setLoading(false);
  }

  const statusConfig = {
    consistent: { color:"#22C55E", label:"CONSISTENT", bg:"#1a3d1a" },
    active:     { color:"#EF4444", label:"ACTIVE ISSUE", bg:"#3d1a1a" },
    improving:  { color:"#F59E0B", label:"IMPROVING", bg:"#3d2e0a" },
    resolved:   { color:"#3B82F6", label:"RESOLVED", bg:"#1e3a5f" },
  };

  return (
    <div>
      {/* Controls */}
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:10 }}>AI INTELLIGENCE ANALYSIS</div>
        <p style={{ fontFamily:"sans-serif", fontSize:12, color:"#4a6581", marginBottom:14, lineHeight:1.5 }}>Claude scans all op feedback and surfaces patterns — strengths, weaknesses, and trends. Weaknesses are tagged as <span style={{color:"#EF4444"}}>Active</span>, <span style={{color:"#F59E0B"}}>Improving</span>, or <span style={{color:"#3B82F6"}}>Resolved</span> based on whether they appear in recent ops (last 12 months).</p>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
          <select value={selectedCreator} onChange={e=>setSelectedCreator(e.target.value)} style={{ background:"#060f1e", border:"1px solid #1e3a5f", borderRadius:6, padding:"8px 10px", color:"#c8d8e8", fontFamily:"monospace", fontSize:11, outline:"none" }}>
            {activeCreators.map(c=><option key={c} value={c}>{c==="ALL"?"All Creators":c}</option>)}
          </select>
          <button onClick={runAnalysis} disabled={loading||allFeedback.length===0} style={{ padding:"8px 20px", background:"#1e3a5f", border:"1px solid #3B82F6", borderRadius:6, color:"#3B82F6", fontFamily:"monospace", fontWeight:700, fontSize:11, letterSpacing:"0.1em", cursor:loading||allFeedback.length===0?"not-allowed":"pointer" }}>
            {loading?"ANALYSING…":"▶ RUN ANALYSIS"}
          </button>
          {allFeedback.length===0&&<span style={{ fontFamily:"sans-serif", fontSize:11, color:"#EF4444" }}>No feedback data available. Connect a Google Sheet or use the fallback demo data.</span>}
        </div>
      </div>

      {err&&<div style={{ background:"#3d1a1a", border:"1px solid #EF4444", borderRadius:8, padding:14, marginBottom:16, fontFamily:"sans-serif", fontSize:12, color:"#EF4444" }}>{err}</div>}

      {loading&&(
        <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:32, textAlign:"center" }}>
          <div style={{ fontFamily:"monospace", fontSize:12, color:"#3B82F6", letterSpacing:"0.1em", marginBottom:8 }}>SCANNING {allFeedback.length} FEEDBACK ITEMS…</div>
          <div style={{ fontFamily:"sans-serif", fontSize:11, color:"#3d5a7a" }}>Identifying patterns across {ops.length} operations</div>
        </div>
      )}

      {analysis&&(
        <>
                      {/* Summary — no highlights, just the summary text */}
            <div style={{ background:"#060f1e", border:`1px solid ${analysis.trend==="improving"?"#22C55E":analysis.trend==="declining"?"#EF4444":"#3B82F6"}`, borderRadius:8, padding:16, marginBottom:16 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6" }}>OVERALL ASSESSMENT</div>
                <span style={{ fontFamily:"monospace", fontSize:10, fontWeight:700, padding:"3px 10px", borderRadius:4, background:analysis.trend==="improving"?"#1a3d1a":analysis.trend==="declining"?"#3d1a1a":"#1e3a5f", color:analysis.trend==="improving"?"#22C55E":analysis.trend==="declining"?"#EF4444":"#3B82F6", letterSpacing:"0.08em" }}>
                  {(analysis.trend||"stable").toUpperCase()} TREND
                </span>
              </div>
              <p style={{ fontFamily:"sans-serif", fontSize:13, color:"#c8d8e8", lineHeight:1.6, margin:0 }}>{analysis.summary}</p>
            </div>

          {/* Strengths + Weaknesses side by side */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, alignItems:"start" }}>
            {/* LEFT — Strengths */}
            <div>
              <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.12em", color:"#22C55E", marginBottom:10 }}>◆ STRENGTHS ({analysis.strengths?.length||0})</div>
              {analysis.strengths?.map((s,i)=>(
                <div key={i} style={{ background:"#0d1e33", border:"1px solid #1a3d2a", borderRadius:8, padding:"12px 14px", marginBottom:8 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:6 }}>
                    <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:13, color:"#22C55E" }}>{s.title}</div>
                    <span style={{ fontFamily:"monospace", fontSize:9, padding:"2px 8px", borderRadius:3, background:"#1a3d1a", color:"#22C55E", letterSpacing:"0.08em", flexShrink:0, marginLeft:8 }}>CONSISTENT</span>
                  </div>
                  <p style={{ fontFamily:"sans-serif", fontSize:12, color:"#8ab8a8", lineHeight:1.5, margin:0 }}>{s.description}</p>
                  {s.opCount>0&&<div style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a", marginTop:6 }}>Seen in ~{s.opCount} ops</div>}
                </div>
              ))}
            </div>
            {/* RIGHT — Weaknesses */}
            <div>
              <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.12em", color:"#EF4444", marginBottom:10 }}>◆ WEAKNESSES ({analysis.weaknesses?.length||0})</div>
              {analysis.weaknesses?.map((w,i)=>{
                const sc = statusConfig[w.status] || statusConfig.active;
                const quotes = (w.evidenceKeys||[]).map(k=>filteredFeedbackWithText[k]?.text).filter(t=>t&&t.trim()).slice(0,2);
                return (
                  <div key={i} style={{ background:"#0d1e33", border:`1px solid ${sc.color}33`, borderRadius:8, padding:"12px 14px", marginBottom:8 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:6 }}>
                      <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:13, color:sc.color }}>{w.title}</div>
                      <span style={{ fontFamily:"monospace", fontSize:9, padding:"2px 8px", borderRadius:3, background:sc.bg, color:sc.color, letterSpacing:"0.08em", flexShrink:0, marginLeft:8 }}>{sc.label}</span>
                    </div>
                    <p style={{ fontFamily:"sans-serif", fontSize:12, color:"#8a9ab8", lineHeight:1.5, margin:"0 0 6px" }}>{w.description}</p>
                    {quotes.map((q,qi)=>(
                      <div key={qi} style={{ fontFamily:"sans-serif", fontSize:11, color:"#3d5a7a", borderLeft:`2px solid ${sc.color}44`, paddingLeft:8, fontStyle:"italic", marginBottom:4 }}>"{q.length>120?q.slice(0,120)+"…":q}"</div>
                    ))}
                    <div style={{ display:"flex", gap:12, marginTop:6 }}>
                      {w.opCount>0&&<span style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a" }}>~{w.opCount} ops</span>}
                      {w.lastSeen&&<span style={{ fontFamily:"monospace", fontSize:10, color:"#3d5a7a" }}>Last: {w.lastSeen}</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ── ANALYTICS ─────────────────────────────────────────────────────────────────
function AnalyticsView({ ops }) {
  const years = [...new Set(ops.map(o=>o.year))].sort();
  const creatorYearData = years.map(yr=>{
    const row={year:`20${yr}`};
    Object.keys(CREATORS).forEach(name=>{
      const yo=ops.filter(o=>o.year===yr&&o.creator===name);
      row[name]=yo.length>0?parseFloat((yo.reduce((s,o)=>s+o.rating,0)/yo.length).toFixed(2)):null;
    });
    return row;
  });
  const topOps=[...ops].sort((a,b)=>b.rating-a.rating).slice(0,8);
  const bottomOps=[...ops].sort((a,b)=>a.rating-b.rating).slice(0,5);
  const totalByCreator=Object.keys(CREATORS).filter(n=>ops.some(o=>o.creator===n)).map(name=>{
    const yo=ops.filter(o=>o.creator===name);
    return{name,ops:yo.length,responses:yo.reduce((s,o)=>s+o.responses,0)};
  });
  return (
    <div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>CREATOR AVG RATING BY YEAR</div>
        <ResponsiveContainer width="100%" height={190}><LineChart data={creatorYearData}><CartesianGrid strokeDasharray="3 3" stroke="#1a2d42"/><XAxis dataKey="year" tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false}/><YAxis domain={[4,10]} tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false} width={22}/><Tooltip contentStyle={{background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:6,fontFamily:"monospace"}} itemStyle={{fontSize:11}}/>{Object.entries(CREATORS).filter(([n])=>ops.some(o=>o.creator===n)).map(([name,style])=><Line key={name} type="monotone" dataKey={name} stroke={style.color} strokeWidth={2} dot={{fill:style.color,r:3}} connectNulls/>)}</LineChart></ResponsiveContainer>
        <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginTop:8 }}>
          {Object.entries(CREATORS).filter(([n])=>ops.some(o=>o.creator===n)).map(([name,style])=>(
            <div key={name} style={{ display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:8, height:8, borderRadius:"50%", background:style.color }}/>
              <span style={{ fontFamily:"sans-serif", fontSize:10, color:"#5b7fa6" }}>{name}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>OPS & RESPONSES BY CREATOR</div>
        <ResponsiveContainer width="100%" height={150}><BarChart data={totalByCreator} barCategoryGap="30%"><CartesianGrid strokeDasharray="3 3" stroke="#1a2d42" vertical={false}/><XAxis dataKey="name" tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"sans-serif"}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#5b7fa6",fontSize:10,fontFamily:"monospace"}} axisLine={false} tickLine={false} width={22}/><Tooltip contentStyle={{background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:6,fontFamily:"monospace"}}/><Bar dataKey="ops" name="Ops" fill="#3B82F6" radius={[3,3,0,0]} barSize={18}/><Bar dataKey="responses" name="Responses" fill="#F59E0B" radius={[3,3,0,0]} barSize={18}/></BarChart></ResponsiveContainer>
      </div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16, marginBottom:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>TOP 8 OPERATIONS</div>
        {topOps.map((op,i)=>(
          <div key={op.name} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
            <div style={{ width:18, fontFamily:"monospace", fontSize:10, color:i<3?"#F59E0B":"#3d5a7a", textAlign:"right" }}>#{i+1}</div>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                <span style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:"#c8d8e8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{op.name}</span>
                <span style={{ fontFamily:"monospace", fontSize:11, fontWeight:700, color:ratingColor(op.rating), flexShrink:0, marginLeft:8 }}>{op.rating.toFixed(2)}</span>
              </div>
              <RatingBar value={op.rating}/>
              <div style={{ fontFamily:"sans-serif", fontSize:10, color:CREATORS[op.creator]?.color||"#5b7fa6", marginTop:2 }}>{op.creator}</div>
            </div>
          </div>
        ))}
      </div>
      <div style={{ background:"#0d1e33", border:"1px solid #1e3a5f", borderRadius:8, padding:16 }}>
        <div style={{ fontFamily:"monospace", fontSize:10, letterSpacing:"0.1em", color:"#5b7fa6", marginBottom:12 }}>IMPROVEMENT TARGETS</div>
        {bottomOps.map(op=>(
          <div key={op.name} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"7px 0", borderBottom:"1px solid #0e1f30" }}>
            <div>
              <div style={{ fontFamily:"sans-serif", fontWeight:700, fontSize:12, color:"#c8d8e8" }}>{op.name}</div>
              <div style={{ fontFamily:"sans-serif", fontSize:10, color:CREATORS[op.creator]?.color||"#5b7fa6" }}>{op.creator}</div>
            </div>
            <RatingBadge value={op.rating}/>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── APP SHELL ─────────────────────────────────────────────────────────────────
export default function App() {
  const [view, setView] = useState("dashboard");
  const [ops, setOps] = useState(FALLBACK_OPS);
  const [fileName, setFileName] = useState(null);
  const [uploadErr, setUploadErr] = useState("");
  const [uploading, setUploading] = useState(false);
  const [feedbackOp, setFeedbackOp] = useState(null);
  const fileRef = useRef(null);

  // Load SheetJS
  useEffect(() => {
    if (window.XLSX) return;
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    document.head.appendChild(s);
  }, []);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true); setUploadErr("");
    try {
      if (!window.XLSX) throw new Error("SheetJS not loaded yet, try again in a moment");
      const data = await parseXLSX(file);
      if (!data.length) throw new Error("No valid ops found. Check your column headers.");
      setOps(data);
      setFileName(file.name);
    } catch(err) { setUploadErr(err.message); console.error("Upload error:", err); }
    setUploading(false);
    e.target.value = "";
  }

  function handleSelectOp(op) { setFeedbackOp(op); setView("feedback"); }

  const VIEWS = [
    { id:"dashboard", label:"OVERVIEW", icon:"◈" },
    { id:"ops", label:"OPS", icon:"⊕" },
    { id:"feedback", label:"FEEDBACK", icon:"↩" },
    { id:"creators", label:"ZEUS", icon:"⚡" },
    { id:"intel", label:"INTEL", icon:"◉" },
    { id:"analytics", label:"ANALYTICS", icon:"▦" },
  ];

  return (
    <div style={{ minHeight:"100vh", background:"#0a1628", color:"#c8d8e8" }}>
      <style>{`*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:#0a1628}::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:2px}select option{background:#0d1e33;color:#c8d8e8}input::placeholder{color:#3d5a7a}`}</style>

      {/* Header */}
      <div style={{ position:"sticky", top:0, zIndex:100, background:"#060f1e", borderBottom:"1px solid #1e3a5f", padding:"0 12px", display:"flex", alignItems:"center", height:50, gap:10 }}>
        <img src={LOGO} alt="AZO" style={{ width:32, height:32, borderRadius:"50%", objectFit:"cover" }} onError={e=>{e.target.style.display="none"}}/>
        <div>
          <div style={{ fontFamily:"monospace", fontWeight:700, fontSize:13, letterSpacing:"0.14em", color:"#e8edf2", lineHeight:1 }}>AUSTRALIAN ZEUS OPS</div>
          <div style={{ fontFamily:"monospace", fontSize:9, color:"#3d5a7a", letterSpacing:"0.08em" }}>OPERATION INTELLIGENCE SYSTEM</div>
        </div>
        <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:8 }}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }}/>
          <button onClick={()=>fileRef.current?.click()} disabled={uploading} style={{ padding:"5px 12px", background:"#1e3a5f", border:"1px solid #3B82F6", borderRadius:5, color:"#3B82F6", fontFamily:"monospace", fontWeight:700, fontSize:9, letterSpacing:"0.1em", cursor:"pointer" }}>
            {uploading ? "LOADING…" : "⬆ UPLOAD XLSX"}
          </button>
        </div>
      </div>

      {/* Nav */}
      <div style={{ position:"sticky", top:50, zIndex:99, background:"#060f1e", borderBottom:"1px solid #0e1f30", padding:"0 8px", display:"flex", gap:0, overflowX:"auto" }}>
        {VIEWS.map(v=>(
          <button key={v.id} onClick={()=>setView(v.id)} style={{ padding:"9px 12px", border:"none", background:"none", cursor:"pointer", fontFamily:"monospace", fontWeight:700, fontSize:10, letterSpacing:"0.1em", color:view===v.id?"#3B82F6":"#3d5a7a", borderBottom:view===v.id?"2px solid #3B82F6":"2px solid transparent", whiteSpace:"nowrap" }}>
            {v.icon} {v.label}
          </button>
        ))}
      </div>

      {/* Status bar */}
      <div style={{ background:"#060f1e", borderBottom:"1px solid #0e1f30", padding:"4px 16px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        {fileName
          ? <span style={{ fontFamily:"monospace", fontSize:9, color:"#22C55E" }}>📊 {fileName} · {ops.length} ops loaded</span>
          : <span style={{ fontFamily:"monospace", fontSize:9, color:"#4a6581" }}>📋 Demo data · Upload your XLSX to load real op data</span>
        }
        {uploadErr && <span style={{ fontFamily:"monospace", fontSize:9, color:"#EF4444" }}>⚠ {uploadErr}</span>}
      </div>

      <div style={{ padding:12, maxWidth:740, margin:"0 auto", paddingBottom:40 }}>
        {view==="dashboard"&&<DashboardView ops={ops}/>}
        {view==="ops"&&<OperationsView ops={ops} onSelectOp={handleSelectOp}/>}
        {view==="feedback"&&<FeedbackView ops={ops} selectedOp={feedbackOp}/>}
        {view==="creators"&&<CreatorsView ops={ops}/>}
        {view==="intel"&&<IntelView ops={ops}/>}
        {view==="analytics"&&<AnalyticsView ops={ops}/>}
      </div>
    </div>
  );
}