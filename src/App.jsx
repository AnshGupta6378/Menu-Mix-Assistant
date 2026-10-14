import React, { useState, useMemo, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

// ============ EMBEDDED MOCK DATA (synthetic - no real restaurant values) ============
const DATA = {"restaurants":[{"id":"MOCK1010","name":"Green Grill","subzone":"Lakeview","cuisine":"biryani","monthly":{"2026-02":{"orders":3635,"revenue":812895,"billsub":948735},"2026-03":{"orders":3871,"revenue":1169492,"billsub":1467109},"2026-04":{"orders":3477,"revenue":1004123,"billsub":1154364},"2026-05":{"orders":3626,"revenue":1247270,"billsub":1464904}},"item_count":101,"items":[{"name":"Paneer Tikka","cat":"Starters","price":380,"orders":796,"rating":4.5,"fc":34},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":300,"orders":781,"rating":4.45,"fc":34},{"name":"Butter Roti","cat":"Breads","price":50,"orders":741,"rating":4.51,"fc":13},{"name":"Tandoori Roti","cat":"Breads","price":40,"orders":726,"rating":4.3,"fc":12},{"name":"Butter Naan","cat":"Breads","price":50,"orders":726,"rating":4.02,"fc":16},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":270,"orders":711,"rating":3.68,"fc":27},{"name":"Kadhai Paneer","cat":"Main Course","price":270,"orders":701,"rating":3.77,"fc":33},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":260,"orders":683,"rating":4.09,"fc":20},{"name":"Brownie","cat":"Desserts","price":170,"orders":672,"rating":3.73,"fc":32},{"name":"Veg Kolhapuri","cat":"Main Course","price":270,"orders":664,"rating":3.76,"fc":30},{"name":"Veg Pulao","cat":"Rice & Biryani","price":240,"orders":657,"rating":4.47,"fc":26},{"name":"Jeera Rice","cat":"Rice & Biryani","price":150,"orders":636,"rating":4.26,"fc":22},{"name":"Veg Chowmein","cat":"Asian Mains","price":220,"orders":633,"rating":4.05,"fc":19},{"name":"Veg Spring Roll","cat":"Starters","price":180,"orders":625,"rating":4.25,"fc":24},{"name":"Chilli Paneer","cat":"Asian Mains","price":330,"orders":618,"rating":4.24,"fc":26},{"name":"Veg Biryani","cat":"Rice & Biryani","price":230,"orders":611,"rating":4.28,"fc":30},{"name":"Gulab Jamun","cat":"Desserts","price":120,"orders":592,"rating":4.44,"fc":28},{"name":"Mushroom Tikka","cat":"Starters","price":290,"orders":513,"rating":4.18,"fc":30},{"name":"Paneer Butter Masala","cat":"Main Course","price":280,"orders":488,"rating":4.59,"fc":32},{"name":"Crispy Corn","cat":"Starters","price":190,"orders":417,"rating":4.43,"fc":22},{"name":"Cheese Naan","cat":"Breads","price":110,"orders":411,"rating":3.91,"fc":22},{"name":"Palak Paneer","cat":"Main Course","price":320,"orders":391,"rating":4.58,"fc":31},{"name":"Garlic Naan","cat":"Breads","price":90,"orders":389,"rating":4.55,"fc":18},{"name":"Veg Manchurian","cat":"Asian Mains","price":290,"orders":370,"rating":4.46,"fc":22},{"name":"Hara Bhara Kabab","cat":"Starters","price":250,"orders":357,"rating":4.06,"fc":28}],"mBillSub":1258778,"mCV":1058445,"mDiscount":144625,"mOrders":3652},{"id":"MOCK1015","name":"Saffron Grill","subzone":"Hillcrest","cuisine":"north indian","monthly":{"2026-02":{"orders":3427,"revenue":1035166,"billsub":1291979},"2026-03":{"orders":3575,"revenue":681590,"billsub":800800},"2026-04":{"orders":3033,"revenue":982514,"billsub":1128276},"2026-05":{"orders":3385,"revenue":1071203,"billsub":1354000}},"item_count":76,"items":[{"name":"Chana Masala","cat":"Main Course","price":200,"orders":780,"rating":4.17,"fc":24},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":330,"orders":774,"rating":4.18,"fc":27},{"name":"Cold Coffee","cat":"Beverages","price":150,"orders":760,"rating":4.37,"fc":20},{"name":"Paneer Tikka","cat":"Starters","price":320,"orders":757,"rating":3.66,"fc":34},{"name":"Masala Chaas","cat":"Beverages","price":50,"orders":742,"rating":4.1,"fc":14},{"name":"Veg Seekh","cat":"Starters","price":290,"orders":696,"rating":4.58,"fc":30},{"name":"Sweet Lassi","cat":"Beverages","price":100,"orders":694,"rating":4.52,"fc":18},{"name":"Mushroom Tikka","cat":"Starters","price":310,"orders":632,"rating":3.67,"fc":30},{"name":"Chilli Paneer","cat":"Asian Mains","price":240,"orders":600,"rating":4.5,"fc":26},{"name":"Gulab Jamun","cat":"Desserts","price":140,"orders":590,"rating":4.4,"fc":28},{"name":"Chilli Mushroom","cat":"Asian Mains","price":240,"orders":589,"rating":3.75,"fc":24},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":220,"orders":572,"rating":4.04,"fc":21},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":190,"orders":562,"rating":4.05,"fc":20},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":505,"rating":4.35,"fc":24},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":474,"rating":4.13,"fc":20},{"name":"Gajar Halwa","cat":"Desserts","price":170,"orders":465,"rating":4.3,"fc":30},{"name":"Jeera Rice","cat":"Rice & Biryani","price":190,"orders":442,"rating":4.13,"fc":22},{"name":"Dal Makhani","cat":"Main Course","price":230,"orders":408,"rating":4.29,"fc":26},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":350,"orders":396,"rating":3.72,"fc":34},{"name":"Veg Pulao","cat":"Rice & Biryani","price":220,"orders":345,"rating":4.06,"fc":26},{"name":"Veg Spring Roll","cat":"Starters","price":180,"orders":333,"rating":4.01,"fc":24},{"name":"Malai Kofta","cat":"Main Course","price":290,"orders":313,"rating":4.55,"fc":34},{"name":"Brownie","cat":"Desserts","price":170,"orders":293,"rating":4.28,"fc":32},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":340,"orders":293,"rating":4.06,"fc":32},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":240,"orders":268,"rating":4.43,"fc":24}],"mBillSub":1143764,"mCV":942618,"mDiscount":151533,"mOrders":3355},{"id":"MOCK1016","name":"Royal Grill","subzone":"Westgate","cuisine":"south indian","monthly":{"2026-02":{"orders":3533,"revenue":871248,"billsub":1112895},"2026-03":{"orders":3577,"revenue":786724,"billsub":1008714},"2026-04":{"orders":3409,"revenue":1070961,"billsub":1346555},"2026-05":{"orders":3304,"revenue":1022621,"billsub":1275344}},"item_count":56,"items":[{"name":"Steamed Rice","cat":"Rice & Biryani","price":110,"orders":781,"rating":4.43,"fc":18},{"name":"Paneer Butter Masala","cat":"Main Course","price":340,"orders":779,"rating":4.07,"fc":32},{"name":"Cheese Naan","cat":"Breads","price":90,"orders":779,"rating":3.87,"fc":22},{"name":"Rava Dosa","cat":"South Indian","price":150,"orders":764,"rating":3.75,"fc":27},{"name":"Hara Bhara Kabab","cat":"Starters","price":260,"orders":741,"rating":3.94,"fc":28},{"name":"Medu Vada","cat":"South Indian","price":110,"orders":724,"rating":3.65,"fc":22},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":290,"orders":698,"rating":4.27,"fc":27},{"name":"Mix Veg","cat":"Main Course","price":230,"orders":616,"rating":4.46,"fc":28},{"name":"Ice Cream","cat":"Desserts","price":120,"orders":615,"rating":4.01,"fc":24},{"name":"Veg Chowmein","cat":"Asian Mains","price":190,"orders":610,"rating":4.55,"fc":19},{"name":"Plain Dosa","cat":"South Indian","price":100,"orders":595,"rating":4.59,"fc":22},{"name":"Gajar Halwa","cat":"Desserts","price":160,"orders":582,"rating":3.73,"fc":30},{"name":"Idli Sambar","cat":"South Indian","price":130,"orders":576,"rating":3.63,"fc":20},{"name":"Malai Kofta","cat":"Main Course","price":280,"orders":549,"rating":4.07,"fc":34},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":220,"orders":547,"rating":3.65,"fc":20},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":540,"rating":4.57,"fc":24},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":320,"orders":524,"rating":3.84,"fc":24},{"name":"Uttapam","cat":"South Indian","price":140,"orders":506,"rating":3.97,"fc":26},{"name":"Butter Naan","cat":"Breads","price":70,"orders":480,"rating":4.41,"fc":16},{"name":"Palak Paneer","cat":"Main Course","price":310,"orders":446,"rating":4.52,"fc":31},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":280,"orders":426,"rating":4.27,"fc":34},{"name":"Kadhai Paneer","cat":"Main Course","price":320,"orders":407,"rating":3.91,"fc":33},{"name":"Garlic Naan","cat":"Breads","price":80,"orders":402,"rating":3.87,"fc":18},{"name":"Paneer Tikka","cat":"Starters","price":290,"orders":377,"rating":3.83,"fc":34},{"name":"Chilli Mushroom","cat":"Asian Mains","price":290,"orders":366,"rating":4.3,"fc":24}],"mBillSub":1185877,"mCV":937888,"mDiscount":198625,"mOrders":3456},{"id":"MOCK1005","name":"Velvet Diner","subzone":"Sunpark","cuisine":"chinese","monthly":{"2026-02":{"orders":3524,"revenue":859356,"billsub":1032532},"2026-03":{"orders":3289,"revenue":788818,"billsub":976833},"2026-04":{"orders":3603,"revenue":1031540,"billsub":1253844},"2026-05":{"orders":3269,"revenue":1061083,"billsub":1219337}},"item_count":100,"items":[{"name":"Malai Kofta","cat":"Main Course","price":280,"orders":800,"rating":4.14,"fc":34},{"name":"Masala Dosa","cat":"South Indian","price":130,"orders":748,"rating":3.84,"fc":25},{"name":"Crispy Corn","cat":"Starters","price":220,"orders":745,"rating":4.19,"fc":22},{"name":"Butter Naan","cat":"Breads","price":60,"orders":703,"rating":3.75,"fc":16},{"name":"Cheese Naan","cat":"Breads","price":130,"orders":680,"rating":4.19,"fc":22},{"name":"Gulab Jamun","cat":"Desserts","price":130,"orders":673,"rating":3.83,"fc":28},{"name":"Sweet Lassi","cat":"Beverages","price":90,"orders":664,"rating":4.27,"fc":18},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":642,"rating":4.59,"fc":24},{"name":"Veg Kolhapuri","cat":"Main Course","price":310,"orders":577,"rating":4.2,"fc":30},{"name":"Fresh Lime Soda","cat":"Beverages","price":70,"orders":546,"rating":4.45,"fc":12},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":260,"orders":543,"rating":3.67,"fc":27},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":240,"orders":542,"rating":3.7,"fc":20},{"name":"Rava Dosa","cat":"South Indian","price":180,"orders":540,"rating":4.16,"fc":27},{"name":"Veg Chowmein","cat":"Asian Mains","price":190,"orders":529,"rating":3.71,"fc":19},{"name":"Uttapam","cat":"South Indian","price":140,"orders":527,"rating":4.25,"fc":26},{"name":"Mix Veg","cat":"Main Course","price":220,"orders":504,"rating":3.87,"fc":28},{"name":"Cold Coffee","cat":"Beverages","price":110,"orders":501,"rating":3.84,"fc":20},{"name":"Rasmalai","cat":"Desserts","price":140,"orders":501,"rating":4.51,"fc":30},{"name":"Butter Roti","cat":"Breads","price":40,"orders":354,"rating":4.2,"fc":13},{"name":"Chilli Mushroom","cat":"Asian Mains","price":290,"orders":351,"rating":3.69,"fc":24},{"name":"Dal Makhani","cat":"Main Course","price":230,"orders":349,"rating":3.61,"fc":26},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":250,"orders":331,"rating":4.07,"fc":24},{"name":"Kadhai Paneer","cat":"Main Course","price":340,"orders":323,"rating":4.32,"fc":33},{"name":"Mango Lassi","cat":"Beverages","price":120,"orders":291,"rating":4.43,"fc":22},{"name":"Hara Bhara Kabab","cat":"Starters","price":230,"orders":277,"rating":3.89,"fc":28}],"mBillSub":1120636,"mCV":935199,"mDiscount":136216,"mOrders":3421},{"id":"MOCK1012","name":"Maple House","subzone":"Riverside","cuisine":"chinese","monthly":{"2026-02":{"orders":3407,"revenue":851789,"billsub":988030},"2026-03":{"orders":2999,"revenue":923645,"billsub":1067644},"2026-04":{"orders":3210,"revenue":934325,"billsub":1190910},"2026-05":{"orders":2906,"revenue":918568,"billsub":1127528}},"item_count":122,"items":[{"name":"Rasmalai","cat":"Desserts","price":160,"orders":784,"rating":4.02,"fc":30},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":310,"orders":766,"rating":3.83,"fc":34},{"name":"Steamed Rice","cat":"Rice & Biryani","price":120,"orders":736,"rating":4.12,"fc":18},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":230,"orders":705,"rating":3.88,"fc":21},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":270,"orders":666,"rating":3.89,"fc":24},{"name":"Masala Dosa","cat":"South Indian","price":130,"orders":662,"rating":4.28,"fc":25},{"name":"Curd Rice","cat":"Rice & Biryani","price":130,"orders":604,"rating":4.24,"fc":20},{"name":"Mango Lassi","cat":"Beverages","price":130,"orders":603,"rating":4.14,"fc":22},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":350,"orders":571,"rating":4.49,"fc":32},{"name":"Mushroom Tikka","cat":"Starters","price":260,"orders":545,"rating":3.63,"fc":30},{"name":"Chilli Paneer","cat":"Asian Mains","price":330,"orders":542,"rating":3.8,"fc":26},{"name":"Gajar Halwa","cat":"Desserts","price":150,"orders":496,"rating":4.2,"fc":30},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":495,"rating":4.6,"fc":14},{"name":"Idli Sambar","cat":"South Indian","price":120,"orders":485,"rating":3.87,"fc":20},{"name":"Veg Manchurian","cat":"Asian Mains","price":250,"orders":436,"rating":4.51,"fc":22},{"name":"Veg Chowmein","cat":"Asian Mains","price":240,"orders":420,"rating":4.04,"fc":19},{"name":"Crispy Corn","cat":"Starters","price":230,"orders":417,"rating":4.07,"fc":22},{"name":"Medu Vada","cat":"South Indian","price":110,"orders":412,"rating":4.11,"fc":22},{"name":"Cold Coffee","cat":"Beverages","price":150,"orders":367,"rating":4.27,"fc":20},{"name":"Paneer Tikka","cat":"Starters","price":280,"orders":356,"rating":4.0,"fc":34},{"name":"Veg Biryani","cat":"Rice & Biryani","price":250,"orders":351,"rating":3.76,"fc":30},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":270,"orders":306,"rating":4.3,"fc":27},{"name":"Jeera Rice","cat":"Rice & Biryani","price":180,"orders":296,"rating":3.8,"fc":22},{"name":"Veg Spring Roll","cat":"Starters","price":160,"orders":245,"rating":4.39,"fc":24},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":224,"rating":4.15,"fc":24}],"mBillSub":1093528,"mCV":907082,"mDiscount":138704,"mOrders":3130},{"id":"MOCK1000","name":"Golden Kitchen","subzone":"Lakeview","cuisine":"south indian","monthly":{"2026-02":{"orders":2930,"revenue":739802,"billsub":922950},"2026-03":{"orders":3142,"revenue":1035253,"billsub":1316498},"2026-04":{"orders":3378,"revenue":1062434,"billsub":1259994},"2026-05":{"orders":2818,"revenue":695039,"billsub":879216}},"item_count":84,"items":[{"name":"Brownie","cat":"Desserts","price":130,"orders":779,"rating":4.48,"fc":32},{"name":"Paneer Tikka","cat":"Starters","price":330,"orders":773,"rating":4.25,"fc":34},{"name":"Rava Dosa","cat":"South Indian","price":140,"orders":717,"rating":3.83,"fc":27},{"name":"Veg Biryani","cat":"Rice & Biryani","price":230,"orders":711,"rating":4.1,"fc":30},{"name":"Jeera Rice","cat":"Rice & Biryani","price":190,"orders":695,"rating":4.29,"fc":22},{"name":"Rasmalai","cat":"Desserts","price":130,"orders":694,"rating":4.11,"fc":30},{"name":"Masala Dosa","cat":"South Indian","price":150,"orders":683,"rating":4.22,"fc":25},{"name":"Plain Dosa","cat":"South Indian","price":130,"orders":631,"rating":3.79,"fc":22},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":200,"orders":620,"rating":4.48,"fc":24},{"name":"Ice Cream","cat":"Desserts","price":90,"orders":552,"rating":4.36,"fc":24},{"name":"Butter Roti","cat":"Breads","price":40,"orders":519,"rating":4.13,"fc":13},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":280,"orders":513,"rating":3.98,"fc":32},{"name":"Gulab Jamun","cat":"Desserts","price":110,"orders":485,"rating":3.76,"fc":28},{"name":"Hara Bhara Kabab","cat":"Starters","price":220,"orders":472,"rating":3.94,"fc":28},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":450,"rating":3.87,"fc":20},{"name":"Tandoori Roti","cat":"Breads","price":40,"orders":448,"rating":3.96,"fc":12},{"name":"Uttapam","cat":"South Indian","price":170,"orders":429,"rating":3.88,"fc":26},{"name":"Veg Spring Roll","cat":"Starters","price":180,"orders":407,"rating":4.45,"fc":24},{"name":"Crispy Corn","cat":"Starters","price":200,"orders":384,"rating":3.7,"fc":22},{"name":"Steamed Rice","cat":"Rice & Biryani","price":110,"orders":372,"rating":4.44,"fc":18},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":360,"orders":274,"rating":4.42,"fc":34},{"name":"Mushroom Tikka","cat":"Starters","price":290,"orders":265,"rating":4.05,"fc":30},{"name":"Veg Pulao","cat":"Rice & Biryani","price":210,"orders":215,"rating":4.13,"fc":26},{"name":"Butter Naan","cat":"Breads","price":50,"orders":203,"rating":4.39,"fc":16},{"name":"Garlic Naan","cat":"Breads","price":100,"orders":181,"rating":4.11,"fc":18}],"mBillSub":1094664,"mCV":883132,"mDiscount":165052,"mOrders":3067},{"id":"MOCK1014","name":"Silver Cafe","subzone":"Hillcrest","cuisine":"fast food","monthly":{"2026-02":{"orders":3338,"revenue":1128158,"billsub":1358566},"2026-03":{"orders":3287,"revenue":765873,"billsub":877629},"2026-04":{"orders":3343,"revenue":808732,"billsub":986185},"2026-05":{"orders":3633,"revenue":782596,"billsub":959112}},"item_count":126,"items":[{"name":"Paneer Biryani","cat":"Rice & Biryani","price":280,"orders":796,"rating":4.32,"fc":32},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":782,"rating":4.47,"fc":14},{"name":"Lachha Paratha","cat":"Breads","price":70,"orders":778,"rating":4.28,"fc":17},{"name":"Mushroom Tikka","cat":"Starters","price":270,"orders":778,"rating":3.87,"fc":30},{"name":"Hara Bhara Kabab","cat":"Starters","price":210,"orders":732,"rating":4.14,"fc":28},{"name":"Malai Kofta","cat":"Main Course","price":280,"orders":696,"rating":4.14,"fc":34},{"name":"Cold Coffee","cat":"Beverages","price":140,"orders":651,"rating":4.07,"fc":20},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":250,"orders":640,"rating":3.84,"fc":21},{"name":"Sweet Lassi","cat":"Beverages","price":100,"orders":637,"rating":4.08,"fc":18},{"name":"Fresh Lime Soda","cat":"Beverages","price":80,"orders":632,"rating":4.18,"fc":12},{"name":"Plain Dosa","cat":"South Indian","price":110,"orders":614,"rating":3.93,"fc":22},{"name":"Chilli Paneer","cat":"Asian Mains","price":260,"orders":591,"rating":3.87,"fc":26},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":260,"orders":583,"rating":3.83,"fc":27},{"name":"Uttapam","cat":"South Indian","price":160,"orders":578,"rating":3.98,"fc":26},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":220,"orders":547,"rating":3.79,"fc":24},{"name":"Steamed Rice","cat":"Rice & Biryani","price":130,"orders":527,"rating":4.08,"fc":18},{"name":"Masala Dosa","cat":"South Indian","price":160,"orders":522,"rating":4.47,"fc":25},{"name":"Crispy Corn","cat":"Starters","price":260,"orders":444,"rating":4.42,"fc":22},{"name":"Idli Sambar","cat":"South Indian","price":90,"orders":361,"rating":4.02,"fc":20},{"name":"Cheese Naan","cat":"Breads","price":100,"orders":344,"rating":4.37,"fc":22},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":336,"rating":4.48,"fc":20},{"name":"Veg Biryani","cat":"Rice & Biryani","price":220,"orders":334,"rating":3.92,"fc":30},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":230,"orders":318,"rating":3.74,"fc":20},{"name":"Butter Roti","cat":"Breads","price":50,"orders":310,"rating":4.21,"fc":13},{"name":"Garlic Naan","cat":"Breads","price":70,"orders":292,"rating":3.88,"fc":18}],"mBillSub":1045373,"mCV":871340,"mDiscount":128173,"mOrders":3400},{"id":"MOCK1018","name":"Silver Table","subzone":"Greenfield","cuisine":"multi cuisine","monthly":{"2026-02":{"orders":3481,"revenue":810623,"billsub":1030376},"2026-03":{"orders":3425,"revenue":1156348,"billsub":1342600},"2026-04":{"orders":3107,"revenue":798937,"billsub":925886},"2026-05":{"orders":3279,"revenue":692921,"billsub":839424}},"item_count":98,"items":[{"name":"Mushroom Tikka","cat":"Starters","price":290,"orders":784,"rating":4.06,"fc":30},{"name":"Uttapam","cat":"South Indian","price":130,"orders":779,"rating":4.3,"fc":26},{"name":"Veg Chowmein","cat":"Asian Mains","price":210,"orders":748,"rating":4.02,"fc":19},{"name":"Veg Manchurian","cat":"Asian Mains","price":230,"orders":702,"rating":4.21,"fc":22},{"name":"Garlic Naan","cat":"Breads","price":90,"orders":672,"rating":3.77,"fc":18},{"name":"Lachha Paratha","cat":"Breads","price":80,"orders":651,"rating":3.99,"fc":17},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":250,"orders":648,"rating":4.24,"fc":20},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":642,"rating":4.29,"fc":14},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":240,"orders":636,"rating":4.0,"fc":21},{"name":"Fresh Lime Soda","cat":"Beverages","price":70,"orders":634,"rating":4.29,"fc":12},{"name":"Crispy Corn","cat":"Starters","price":220,"orders":617,"rating":3.8,"fc":22},{"name":"Butter Naan","cat":"Breads","price":70,"orders":589,"rating":3.79,"fc":16},{"name":"Butter Roti","cat":"Breads","price":40,"orders":541,"rating":4.13,"fc":13},{"name":"Masala Dosa","cat":"South Indian","price":130,"orders":534,"rating":3.95,"fc":25},{"name":"Tandoori Roti","cat":"Breads","price":30,"orders":508,"rating":3.71,"fc":12},{"name":"Rasmalai","cat":"Desserts","price":160,"orders":500,"rating":4.19,"fc":30},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":488,"rating":4.4,"fc":24},{"name":"Sweet Lassi","cat":"Beverages","price":80,"orders":423,"rating":3.99,"fc":18},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":250,"orders":408,"rating":4.3,"fc":24},{"name":"Plain Dosa","cat":"South Indian","price":110,"orders":391,"rating":4.06,"fc":22},{"name":"Gulab Jamun","cat":"Desserts","price":130,"orders":378,"rating":4.16,"fc":28},{"name":"Medu Vada","cat":"South Indian","price":100,"orders":347,"rating":3.64,"fc":22},{"name":"Chilli Mushroom","cat":"Asian Mains","price":230,"orders":294,"rating":4.26,"fc":24},{"name":"Gajar Halwa","cat":"Desserts","price":190,"orders":277,"rating":4.4,"fc":30},{"name":"Cold Coffee","cat":"Beverages","price":110,"orders":267,"rating":4.11,"fc":20}],"mBillSub":1034572,"mCV":864707,"mDiscount":124353,"mOrders":3323},{"id":"MOCK1011","name":"Silver Corner","subzone":"Greenfield","cuisine":"fast food","monthly":{"2026-02":{"orders":2725,"revenue":672220,"billsub":814775},"2026-03":{"orders":3106,"revenue":924410,"billsub":1115054},"2026-04":{"orders":2567,"revenue":845827,"billsub":988295},"2026-05":{"orders":2749,"revenue":1001470,"billsub":1151831}},"item_count":128,"items":[{"name":"Paneer Tikka","cat":"Starters","price":350,"orders":752,"rating":4.08,"fc":34},{"name":"Garlic Naan","cat":"Breads","price":80,"orders":696,"rating":4.08,"fc":18},{"name":"Chana Masala","cat":"Main Course","price":230,"orders":683,"rating":3.7,"fc":24},{"name":"Lachha Paratha","cat":"Breads","price":80,"orders":659,"rating":3.74,"fc":17},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":270,"orders":651,"rating":4.14,"fc":21},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":320,"orders":580,"rating":4.02,"fc":24},{"name":"Mix Veg","cat":"Main Course","price":200,"orders":578,"rating":4.14,"fc":28},{"name":"Veg Kolhapuri","cat":"Main Course","price":280,"orders":576,"rating":4.57,"fc":30},{"name":"Dal Makhani","cat":"Main Course","price":270,"orders":559,"rating":4.15,"fc":26},{"name":"Kadhai Paneer","cat":"Main Course","price":310,"orders":521,"rating":3.64,"fc":33},{"name":"Veg Chowmein","cat":"Asian Mains","price":220,"orders":520,"rating":4.05,"fc":19},{"name":"Hara Bhara Kabab","cat":"Starters","price":270,"orders":508,"rating":3.64,"fc":28},{"name":"Butter Naan","cat":"Breads","price":60,"orders":487,"rating":3.69,"fc":16},{"name":"Chilli Mushroom","cat":"Asian Mains","price":230,"orders":469,"rating":4.24,"fc":24},{"name":"Shahi Paneer","cat":"Main Course","price":280,"orders":466,"rating":3.79,"fc":33},{"name":"Rasmalai","cat":"Desserts","price":140,"orders":445,"rating":4.1,"fc":30},{"name":"Palak Paneer","cat":"Main Course","price":280,"orders":422,"rating":3.85,"fc":31},{"name":"Tandoori Roti","cat":"Breads","price":40,"orders":319,"rating":4.05,"fc":12},{"name":"Veg Manchurian","cat":"Asian Mains","price":270,"orders":283,"rating":3.9,"fc":22},{"name":"Mushroom Tikka","cat":"Starters","price":260,"orders":244,"rating":4.49,"fc":30},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":219,"rating":4.43,"fc":24},{"name":"Chilli Paneer","cat":"Asian Mains","price":260,"orders":193,"rating":4.15,"fc":26},{"name":"Crispy Corn","cat":"Starters","price":200,"orders":178,"rating":3.66,"fc":22},{"name":"Butter Roti","cat":"Breads","price":40,"orders":167,"rating":3.89,"fc":13},{"name":"Cheese Naan","cat":"Breads","price":110,"orders":145,"rating":4.31,"fc":22}],"mBillSub":1017489,"mCV":860982,"mDiscount":111192,"mOrders":2787},{"id":"MOCK1009","name":"Green Bistro","subzone":"Lakeview","cuisine":"fast food","monthly":{"2026-02":{"orders":2955,"revenue":767929,"billsub":895365},"2026-03":{"orders":3561,"revenue":953994,"billsub":1200057},"2026-04":{"orders":3543,"revenue":713801,"billsub":857406},"2026-05":{"orders":3011,"revenue":815309,"billsub":996641}},"item_count":112,"items":[{"name":"Chilli Paneer","cat":"Asian Mains","price":250,"orders":796,"rating":3.97,"fc":26},{"name":"Chana Masala","cat":"Main Course","price":210,"orders":760,"rating":3.83,"fc":24},{"name":"Veg Seekh","cat":"Starters","price":280,"orders":730,"rating":3.68,"fc":30},{"name":"Malai Kofta","cat":"Main Course","price":300,"orders":724,"rating":4.41,"fc":34},{"name":"Veg Manchurian","cat":"Asian Mains","price":240,"orders":721,"rating":4.54,"fc":22},{"name":"Butter Roti","cat":"Breads","price":50,"orders":668,"rating":4.27,"fc":13},{"name":"Lachha Paratha","cat":"Breads","price":80,"orders":627,"rating":3.83,"fc":17},{"name":"Gulab Jamun","cat":"Desserts","price":130,"orders":627,"rating":4.0,"fc":28},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":621,"rating":4.2,"fc":24},{"name":"Tandoori Roti","cat":"Breads","price":40,"orders":614,"rating":4.57,"fc":12},{"name":"Garlic Naan","cat":"Breads","price":80,"orders":609,"rating":3.97,"fc":18},{"name":"Cheese Naan","cat":"Breads","price":100,"orders":559,"rating":3.81,"fc":22},{"name":"Kadhai Paneer","cat":"Main Course","price":260,"orders":480,"rating":3.78,"fc":33},{"name":"Idli Sambar","cat":"South Indian","price":100,"orders":477,"rating":3.91,"fc":20},{"name":"Hara Bhara Kabab","cat":"Starters","price":240,"orders":474,"rating":4.09,"fc":28},{"name":"Mango Lassi","cat":"Beverages","price":130,"orders":450,"rating":3.67,"fc":22},{"name":"Masala Dosa","cat":"South Indian","price":160,"orders":448,"rating":4.02,"fc":25},{"name":"Sweet Lassi","cat":"Beverages","price":90,"orders":425,"rating":4.02,"fc":18},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":240,"orders":377,"rating":4.27,"fc":21},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":356,"rating":4.52,"fc":14},{"name":"Paneer Butter Masala","cat":"Main Course","price":340,"orders":336,"rating":3.93,"fc":32},{"name":"Gajar Halwa","cat":"Desserts","price":170,"orders":336,"rating":3.72,"fc":30},{"name":"Mushroom Tikka","cat":"Starters","price":280,"orders":319,"rating":3.82,"fc":30},{"name":"Rava Dosa","cat":"South Indian","price":160,"orders":273,"rating":4.01,"fc":27},{"name":"Butter Naan","cat":"Breads","price":60,"orders":269,"rating":4.45,"fc":16}],"mBillSub":987367,"mCV":812758,"mDiscount":131832,"mOrders":3268},{"id":"MOCK1007","name":"Golden Corner","subzone":"Riverside","cuisine":"fast food","monthly":{"2026-02":{"orders":2960,"revenue":752922,"billsub":929440},"2026-03":{"orders":3311,"revenue":1023593,"billsub":1221759},"2026-04":{"orders":3289,"revenue":894831,"billsub":1036035},"2026-05":{"orders":3030,"revenue":558194,"billsub":684780}},"item_count":116,"items":[{"name":"Jeera Rice","cat":"Rice & Biryani","price":150,"orders":775,"rating":4.49,"fc":22},{"name":"Veg Kolhapuri","cat":"Main Course","price":310,"orders":746,"rating":3.85,"fc":30},{"name":"Mango Lassi","cat":"Beverages","price":110,"orders":732,"rating":4.32,"fc":22},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":330,"orders":715,"rating":4.28,"fc":27},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":190,"orders":686,"rating":4.28,"fc":24},{"name":"Malai Kofta","cat":"Main Course","price":340,"orders":641,"rating":3.9,"fc":34},{"name":"Kadhai Paneer","cat":"Main Course","price":320,"orders":637,"rating":3.74,"fc":33},{"name":"Veg Chowmein","cat":"Asian Mains","price":190,"orders":619,"rating":4.47,"fc":19},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":610,"rating":3.97,"fc":14},{"name":"Steamed Rice","cat":"Rice & Biryani","price":120,"orders":594,"rating":3.64,"fc":18},{"name":"Gajar Halwa","cat":"Desserts","price":140,"orders":592,"rating":4.48,"fc":30},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":586,"rating":3.98,"fc":20},{"name":"Medu Vada","cat":"South Indian","price":90,"orders":583,"rating":4.01,"fc":22},{"name":"Brownie","cat":"Desserts","price":150,"orders":553,"rating":4.51,"fc":32},{"name":"Dal Makhani","cat":"Main Course","price":260,"orders":516,"rating":4.5,"fc":26},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":465,"rating":3.78,"fc":24},{"name":"Idli Sambar","cat":"South Indian","price":100,"orders":407,"rating":3.97,"fc":20},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":270,"orders":400,"rating":3.82,"fc":24},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":240,"orders":333,"rating":4.38,"fc":20},{"name":"Rasmalai","cat":"Desserts","price":160,"orders":303,"rating":4.51,"fc":30},{"name":"Shahi Paneer","cat":"Main Course","price":300,"orders":300,"rating":4.06,"fc":33},{"name":"Veg Manchurian","cat":"Asian Mains","price":220,"orders":296,"rating":4.28,"fc":22},{"name":"Veg Biryani","cat":"Rice & Biryani","price":230,"orders":272,"rating":4.35,"fc":30},{"name":"Palak Paneer","cat":"Main Course","price":320,"orders":243,"rating":4.6,"fc":31},{"name":"Paneer Butter Masala","cat":"Main Course","price":330,"orders":201,"rating":3.67,"fc":32}],"mBillSub":968004,"mCV":807385,"mDiscount":118124,"mOrders":3148},{"id":"MOCK1003","name":"Lotus Kitchen","subzone":"Old Town","cuisine":"biryani","monthly":{"2026-02":{"orders":2992,"revenue":982178,"billsub":1190816},"2026-03":{"orders":2885,"revenue":642420,"billsub":816455},"2026-04":{"orders":2930,"revenue":593788,"billsub":744220},"2026-05":{"orders":3277,"revenue":972377,"billsub":1189551}},"item_count":141,"items":[{"name":"Plain Dosa","cat":"South Indian","price":130,"orders":777,"rating":4.5,"fc":22},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":735,"rating":4.58,"fc":24},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":728,"rating":3.91,"fc":24},{"name":"Masala Dosa","cat":"South Indian","price":140,"orders":722,"rating":4.01,"fc":25},{"name":"Veg Biryani","cat":"Rice & Biryani","price":220,"orders":711,"rating":4.29,"fc":30},{"name":"Steamed Rice","cat":"Rice & Biryani","price":120,"orders":684,"rating":4.18,"fc":18},{"name":"Medu Vada","cat":"South Indian","price":120,"orders":678,"rating":3.91,"fc":22},{"name":"Chana Masala","cat":"Main Course","price":240,"orders":670,"rating":4.17,"fc":24},{"name":"Cold Coffee","cat":"Beverages","price":120,"orders":660,"rating":4.58,"fc":20},{"name":"Mix Veg","cat":"Main Course","price":210,"orders":633,"rating":4.21,"fc":28},{"name":"Sweet Lassi","cat":"Beverages","price":80,"orders":630,"rating":3.79,"fc":18},{"name":"Veg Seekh","cat":"Starters","price":230,"orders":567,"rating":4.26,"fc":30},{"name":"Gulab Jamun","cat":"Desserts","price":110,"orders":511,"rating":3.95,"fc":28},{"name":"Palak Paneer","cat":"Main Course","price":270,"orders":492,"rating":4.04,"fc":31},{"name":"Mango Lassi","cat":"Beverages","price":100,"orders":486,"rating":3.6,"fc":22},{"name":"Brownie","cat":"Desserts","price":160,"orders":448,"rating":3.93,"fc":32},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":433,"rating":4.09,"fc":20},{"name":"Masala Chaas","cat":"Beverages","price":60,"orders":393,"rating":4.33,"fc":14},{"name":"Rava Dosa","cat":"South Indian","price":160,"orders":342,"rating":4.15,"fc":27},{"name":"Gajar Halwa","cat":"Desserts","price":180,"orders":325,"rating":3.95,"fc":30},{"name":"Fresh Lime Soda","cat":"Beverages","price":70,"orders":311,"rating":3.64,"fc":12},{"name":"Rasmalai","cat":"Desserts","price":160,"orders":273,"rating":3.82,"fc":30},{"name":"Crispy Corn","cat":"Starters","price":250,"orders":190,"rating":3.62,"fc":22},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":260,"orders":149,"rating":4.38,"fc":24},{"name":"Hara Bhara Kabab","cat":"Starters","price":260,"orders":135,"rating":4.42,"fc":28}],"mBillSub":985260,"mCV":797691,"mDiscount":145586,"mOrders":3021},{"id":"MOCK1017","name":"Sunrise Table","subzone":"Sunpark","cuisine":"fast food","monthly":{"2026-02":{"orders":2450,"revenue":635498,"billsub":727650},"2026-03":{"orders":2604,"revenue":824766,"billsub":984312},"2026-04":{"orders":2852,"revenue":797449,"billsub":946864},"2026-05":{"orders":2827,"revenue":852734,"billsub":1091222}},"item_count":62,"items":[{"name":"Dal Khichdi","cat":"Rice & Biryani","price":230,"orders":769,"rating":4.42,"fc":24},{"name":"Veg Pulao","cat":"Rice & Biryani","price":180,"orders":748,"rating":4.49,"fc":26},{"name":"Uttapam","cat":"South Indian","price":130,"orders":745,"rating":4.32,"fc":26},{"name":"Garlic Naan","cat":"Breads","price":80,"orders":671,"rating":3.8,"fc":18},{"name":"Paneer Butter Masala","cat":"Main Course","price":310,"orders":651,"rating":4.2,"fc":32},{"name":"Masala Chaas","cat":"Beverages","price":60,"orders":649,"rating":4.0,"fc":14},{"name":"Palak Paneer","cat":"Main Course","price":280,"orders":610,"rating":4.15,"fc":31},{"name":"Tandoori Roti","cat":"Breads","price":30,"orders":535,"rating":4.38,"fc":12},{"name":"Sweet Lassi","cat":"Beverages","price":100,"orders":501,"rating":4.09,"fc":18},{"name":"Malai Kofta","cat":"Main Course","price":340,"orders":495,"rating":3.96,"fc":34},{"name":"Medu Vada","cat":"South Indian","price":100,"orders":438,"rating":4.28,"fc":22},{"name":"Idli Sambar","cat":"South Indian","price":110,"orders":433,"rating":4.36,"fc":20},{"name":"Gajar Halwa","cat":"Desserts","price":150,"orders":429,"rating":3.77,"fc":30},{"name":"Mango Lassi","cat":"Beverages","price":120,"orders":424,"rating":4.03,"fc":22},{"name":"Fresh Lime Soda","cat":"Beverages","price":80,"orders":388,"rating":4.21,"fc":12},{"name":"Veg Biryani","cat":"Rice & Biryani","price":260,"orders":327,"rating":4.5,"fc":30},{"name":"Ice Cream","cat":"Desserts","price":90,"orders":314,"rating":4.09,"fc":24},{"name":"Lachha Paratha","cat":"Breads","price":70,"orders":292,"rating":4.32,"fc":17},{"name":"Gulab Jamun","cat":"Desserts","price":130,"orders":288,"rating":4.17,"fc":28},{"name":"Chana Masala","cat":"Main Course","price":180,"orders":197,"rating":3.98,"fc":24},{"name":"Dal Makhani","cat":"Main Course","price":270,"orders":197,"rating":4.34,"fc":26},{"name":"Kadhai Paneer","cat":"Main Course","price":290,"orders":175,"rating":3.69,"fc":33},{"name":"Shahi Paneer","cat":"Main Course","price":290,"orders":149,"rating":3.69,"fc":33},{"name":"Rasmalai","cat":"Desserts","price":140,"orders":126,"rating":4.32,"fc":30},{"name":"Curd Rice","cat":"Rice & Biryani","price":150,"orders":87,"rating":4.52,"fc":20}],"mBillSub":937512,"mCV":777612,"mDiscount":118973,"mOrders":2683},{"id":"MOCK1006","name":"Maple Cafe","subzone":"Lakeview","cuisine":"north indian","monthly":{"2026-02":{"orders":2908,"revenue":915699,"billsub":1163200},"2026-03":{"orders":2706,"revenue":840763,"billsub":1052634},"2026-04":{"orders":2856,"revenue":609331,"billsub":736848},"2026-05":{"orders":2515,"revenue":684677,"billsub":875220}},"item_count":87,"items":[{"name":"Gajar Halwa","cat":"Desserts","price":150,"orders":762,"rating":3.87,"fc":30},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":350,"orders":754,"rating":3.84,"fc":32},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":240,"orders":718,"rating":4.55,"fc":21},{"name":"Garlic Naan","cat":"Breads","price":70,"orders":685,"rating":4.21,"fc":18},{"name":"Chilli Mushroom","cat":"Asian Mains","price":240,"orders":639,"rating":4.04,"fc":24},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":370,"orders":633,"rating":4.6,"fc":34},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":250,"orders":603,"rating":3.64,"fc":20},{"name":"Uttapam","cat":"South Indian","price":160,"orders":550,"rating":4.23,"fc":26},{"name":"Rasmalai","cat":"Desserts","price":160,"orders":523,"rating":3.84,"fc":30},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":230,"orders":506,"rating":4.53,"fc":24},{"name":"Curd Rice","cat":"Rice & Biryani","price":150,"orders":460,"rating":3.64,"fc":20},{"name":"Butter Roti","cat":"Breads","price":30,"orders":457,"rating":3.69,"fc":13},{"name":"Veg Chowmein","cat":"Asian Mains","price":240,"orders":453,"rating":4.46,"fc":19},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":432,"rating":3.79,"fc":24},{"name":"Steamed Rice","cat":"Rice & Biryani","price":120,"orders":380,"rating":4.15,"fc":18},{"name":"Chilli Paneer","cat":"Asian Mains","price":320,"orders":377,"rating":4.01,"fc":26},{"name":"Veg Manchurian","cat":"Asian Mains","price":240,"orders":368,"rating":4.27,"fc":22},{"name":"Jeera Rice","cat":"Rice & Biryani","price":180,"orders":329,"rating":4.33,"fc":22},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":260,"orders":298,"rating":3.92,"fc":27},{"name":"Plain Dosa","cat":"South Indian","price":120,"orders":248,"rating":3.87,"fc":22},{"name":"Rava Dosa","cat":"South Indian","price":160,"orders":232,"rating":4.12,"fc":27},{"name":"Lachha Paratha","cat":"Breads","price":70,"orders":202,"rating":3.91,"fc":17},{"name":"Veg Biryani","cat":"Rice & Biryani","price":250,"orders":192,"rating":4.05,"fc":30},{"name":"Butter Naan","cat":"Breads","price":70,"orders":156,"rating":4.06,"fc":16},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":330,"orders":133,"rating":4.07,"fc":24}],"mBillSub":956976,"mCV":762618,"mDiscount":154220,"mOrders":2746},{"id":"MOCK1004","name":"Copper Cafe","subzone":"Sunpark","cuisine":"fast food","monthly":{"2026-02":{"orders":2587,"revenue":834823,"billsub":1021865},"2026-03":{"orders":2267,"revenue":659861,"billsub":825188},"2026-04":{"orders":2378,"revenue":634618,"billsub":782362},"2026-05":{"orders":2242,"revenue":726429,"billsub":863170}},"item_count":98,"items":[{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":220,"orders":735,"rating":4.4,"fc":21},{"name":"Medu Vada","cat":"South Indian","price":110,"orders":657,"rating":4.5,"fc":22},{"name":"Veg Chowmein","cat":"Asian Mains","price":200,"orders":617,"rating":4.26,"fc":19},{"name":"Hara Bhara Kabab","cat":"Starters","price":230,"orders":593,"rating":3.68,"fc":28},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":350,"orders":586,"rating":3.63,"fc":24},{"name":"Mango Lassi","cat":"Beverages","price":110,"orders":524,"rating":4.27,"fc":22},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":240,"orders":518,"rating":4.52,"fc":20},{"name":"Crispy Corn","cat":"Starters","price":210,"orders":517,"rating":4.02,"fc":22},{"name":"Veg Manchurian","cat":"Asian Mains","price":210,"orders":478,"rating":3.74,"fc":22},{"name":"Rasmalai","cat":"Desserts","price":170,"orders":471,"rating":3.85,"fc":30},{"name":"Masala Dosa","cat":"South Indian","price":170,"orders":469,"rating":4.14,"fc":25},{"name":"Rava Dosa","cat":"South Indian","price":150,"orders":439,"rating":4.45,"fc":27},{"name":"Veg Spring Roll","cat":"Starters","price":160,"orders":432,"rating":4.29,"fc":24},{"name":"Gajar Halwa","cat":"Desserts","price":180,"orders":428,"rating":3.88,"fc":30},{"name":"Curd Rice","cat":"Rice & Biryani","price":160,"orders":424,"rating":4.0,"fc":20},{"name":"Ice Cream","cat":"Desserts","price":90,"orders":398,"rating":3.82,"fc":24},{"name":"Veg Seekh","cat":"Starters","price":230,"orders":332,"rating":3.83,"fc":30},{"name":"Veg Biryani","cat":"Rice & Biryani","price":270,"orders":288,"rating":3.7,"fc":30},{"name":"Cold Coffee","cat":"Beverages","price":140,"orders":263,"rating":4.07,"fc":20},{"name":"Veg Pulao","cat":"Rice & Biryani","price":220,"orders":207,"rating":4.56,"fc":26},{"name":"Steamed Rice","cat":"Rice & Biryani","price":130,"orders":117,"rating":4.19,"fc":18},{"name":"Mushroom Tikka","cat":"Starters","price":310,"orders":105,"rating":4.01,"fc":30},{"name":"Chilli Paneer","cat":"Asian Mains","price":280,"orders":91,"rating":3.86,"fc":26},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":270,"orders":66,"rating":4.53,"fc":32},{"name":"Masala Chaas","cat":"Beverages","price":60,"orders":60,"rating":4.22,"fc":14}],"mBillSub":873146,"mCV":713933,"mDiscount":121638,"mOrders":2368},{"id":"MOCK1001","name":"Metro Kitchen","subzone":"Greenfield","cuisine":"fast food","monthly":{"2026-02":{"orders":2690,"revenue":897356,"billsub":1084070},"2026-03":{"orders":2495,"revenue":536531,"billsub":686125},"2026-04":{"orders":2703,"revenue":840143,"billsub":967674},"2026-05":{"orders":2561,"revenue":486315,"billsub":594152}},"item_count":106,"items":[{"name":"Sweet Lassi","cat":"Beverages","price":100,"orders":788,"rating":4.6,"fc":18},{"name":"Kadhai Paneer","cat":"Main Course","price":360,"orders":787,"rating":4.29,"fc":33},{"name":"Idli Sambar","cat":"South Indian","price":90,"orders":764,"rating":4.23,"fc":20},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":260,"orders":714,"rating":4.03,"fc":24},{"name":"Schezwan Fried Rice","cat":"Asian Mains","price":200,"orders":709,"rating":4.58,"fc":21},{"name":"Medu Vada","cat":"South Indian","price":120,"orders":624,"rating":4.18,"fc":22},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":280,"orders":603,"rating":3.7,"fc":34},{"name":"Plain Dosa","cat":"South Indian","price":110,"orders":592,"rating":3.73,"fc":22},{"name":"Mango Lassi","cat":"Beverages","price":120,"orders":532,"rating":3.81,"fc":22},{"name":"Paneer Chilli Dry","cat":"Asian Mains","price":300,"orders":518,"rating":4.46,"fc":27},{"name":"Fresh Lime Soda","cat":"Beverages","price":80,"orders":505,"rating":3.89,"fc":12},{"name":"Malai Kofta","cat":"Main Course","price":310,"orders":488,"rating":4.5,"fc":34},{"name":"Palak Paneer","cat":"Main Course","price":320,"orders":473,"rating":4.56,"fc":31},{"name":"Chilli Paneer","cat":"Asian Mains","price":250,"orders":452,"rating":4.33,"fc":26},{"name":"Shahi Paneer","cat":"Main Course","price":330,"orders":448,"rating":4.6,"fc":33},{"name":"Cold Coffee","cat":"Beverages","price":110,"orders":439,"rating":3.87,"fc":20},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":280,"orders":295,"rating":4.47,"fc":32},{"name":"Rava Dosa","cat":"South Indian","price":140,"orders":283,"rating":3.88,"fc":27},{"name":"Masala Dosa","cat":"South Indian","price":150,"orders":275,"rating":4.19,"fc":25},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":230,"orders":183,"rating":4.02,"fc":24},{"name":"Masala Chaas","cat":"Beverages","price":60,"orders":100,"rating":3.76,"fc":14},{"name":"Uttapam","cat":"South Indian","price":140,"orders":72,"rating":4.46,"fc":26}],"mBillSub":833005,"mCV":690086,"mDiscount":106598,"mOrders":2612},{"id":"MOCK1019","name":"Urban Cafe","subzone":"Sunpark","cuisine":"multi cuisine","monthly":{"2026-02":{"orders":2574,"revenue":746358,"billsub":885456},"2026-03":{"orders":2380,"revenue":789425,"billsub":994840},"2026-04":{"orders":2428,"revenue":665791,"billsub":818236},"2026-05":{"orders":2748,"revenue":517791,"billsub":648528}},"item_count":95,"items":[{"name":"Paneer Butter Masala","cat":"Main Course","price":290,"orders":760,"rating":3.89,"fc":32},{"name":"Gulab Jamun","cat":"Desserts","price":130,"orders":751,"rating":4.54,"fc":28},{"name":"Veg Kolhapuri","cat":"Main Course","price":310,"orders":711,"rating":3.73,"fc":30},{"name":"Fresh Lime Soda","cat":"Beverages","price":80,"orders":674,"rating":3.65,"fc":12},{"name":"Chana Masala","cat":"Main Course","price":220,"orders":651,"rating":4.14,"fc":24},{"name":"Uttapam","cat":"South Indian","price":170,"orders":619,"rating":4.34,"fc":26},{"name":"Masala Dosa","cat":"South Indian","price":160,"orders":602,"rating":4.11,"fc":25},{"name":"Rava Dosa","cat":"South Indian","price":150,"orders":590,"rating":3.89,"fc":27},{"name":"Shahi Paneer","cat":"Main Course","price":320,"orders":566,"rating":4.59,"fc":33},{"name":"Cold Coffee","cat":"Beverages","price":130,"orders":546,"rating":4.01,"fc":20},{"name":"Tandoori Roti","cat":"Breads","price":40,"orders":500,"rating":3.86,"fc":12},{"name":"Kadhai Paneer","cat":"Main Course","price":340,"orders":461,"rating":4.52,"fc":33},{"name":"Masala Chaas","cat":"Beverages","price":60,"orders":433,"rating":4.56,"fc":14},{"name":"Idli Sambar","cat":"South Indian","price":130,"orders":415,"rating":4.14,"fc":20},{"name":"Gajar Halwa","cat":"Desserts","price":170,"orders":315,"rating":4.23,"fc":30},{"name":"Medu Vada","cat":"South Indian","price":110,"orders":314,"rating":4.38,"fc":22},{"name":"Mix Veg","cat":"Main Course","price":270,"orders":295,"rating":3.75,"fc":28},{"name":"Lachha Paratha","cat":"Breads","price":80,"orders":293,"rating":3.9,"fc":17},{"name":"Plain Dosa","cat":"South Indian","price":120,"orders":265,"rating":3.89,"fc":22},{"name":"Dal Makhani","cat":"Main Course","price":220,"orders":182,"rating":4.38,"fc":26},{"name":"Cheese Naan","cat":"Breads","price":130,"orders":144,"rating":4.34,"fc":22},{"name":"Malai Kofta","cat":"Main Course","price":320,"orders":137,"rating":3.74,"fc":34},{"name":"Brownie","cat":"Desserts","price":170,"orders":135,"rating":4.0,"fc":32},{"name":"Rasmalai","cat":"Desserts","price":140,"orders":117,"rating":4.23,"fc":30},{"name":"Palak Paneer","cat":"Main Course","price":320,"orders":100,"rating":4.17,"fc":31}],"mBillSub":836765,"mCV":679841,"mDiscount":121142,"mOrders":2532},{"id":"MOCK1013","name":"Garden Corner","subzone":"Maple Heights","cuisine":"north indian","monthly":{"2026-02":{"orders":2691,"revenue":654653,"billsub":820755},"2026-03":{"orders":2913,"revenue":613572,"billsub":751554},"2026-04":{"orders":2911,"revenue":692854,"billsub":882033},"2026-05":{"orders":2797,"revenue":698034,"billsub":895040}},"item_count":139,"items":[{"name":"Jeera Rice","cat":"Rice & Biryani","price":180,"orders":791,"rating":4.59,"fc":22},{"name":"Crispy Corn","cat":"Starters","price":250,"orders":784,"rating":3.97,"fc":22},{"name":"Plain Dosa","cat":"South Indian","price":130,"orders":778,"rating":3.91,"fc":22},{"name":"Veg Pulao","cat":"Rice & Biryani","price":200,"orders":745,"rating":3.66,"fc":26},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":741,"rating":4.01,"fc":20},{"name":"Masala Dosa","cat":"South Indian","price":140,"orders":670,"rating":4.02,"fc":25},{"name":"Mushroom Tikka","cat":"Starters","price":300,"orders":665,"rating":3.72,"fc":30},{"name":"Cheese Naan","cat":"Breads","price":100,"orders":637,"rating":3.95,"fc":22},{"name":"Chana Masala","cat":"Main Course","price":250,"orders":611,"rating":4.54,"fc":24},{"name":"Butter Roti","cat":"Breads","price":50,"orders":557,"rating":4.06,"fc":13},{"name":"Lachha Paratha","cat":"Breads","price":80,"orders":549,"rating":3.62,"fc":17},{"name":"Tandoori Roti","cat":"Breads","price":30,"orders":546,"rating":4.2,"fc":12},{"name":"Veg Biryani","cat":"Rice & Biryani","price":240,"orders":535,"rating":4.06,"fc":30},{"name":"Paneer Biryani","cat":"Rice & Biryani","price":290,"orders":448,"rating":3.73,"fc":32},{"name":"Hara Bhara Kabab","cat":"Starters","price":230,"orders":435,"rating":4.1,"fc":28},{"name":"Hyderabadi Biryani","cat":"Rice & Biryani","price":340,"orders":349,"rating":4.08,"fc":34},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":326,"rating":4.43,"fc":24},{"name":"Uttapam","cat":"South Indian","price":150,"orders":313,"rating":4.05,"fc":26},{"name":"Mix Veg","cat":"Main Course","price":260,"orders":243,"rating":4.2,"fc":28},{"name":"Paneer Tikka","cat":"Starters","price":340,"orders":239,"rating":3.96,"fc":34},{"name":"Idli Sambar","cat":"South Indian","price":130,"orders":203,"rating":4.41,"fc":20},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":210,"orders":187,"rating":3.91,"fc":24},{"name":"Rava Dosa","cat":"South Indian","price":140,"orders":91,"rating":3.89,"fc":27},{"name":"Veg Seekh","cat":"Starters","price":260,"orders":86,"rating":4.0,"fc":30},{"name":"Shahi Paneer","cat":"Main Course","price":300,"orders":81,"rating":3.65,"fc":33}],"mBillSub":837346,"mCV":664778,"mDiscount":137578,"mOrders":2828},{"id":"MOCK1008","name":"Garden Grill","subzone":"Greenfield","cuisine":"biryani","monthly":{"2026-02":{"orders":2106,"revenue":415064,"billsub":478062},"2026-03":{"orders":2115,"revenue":567642,"billsub":681030},"2026-04":{"orders":1957,"revenue":619599,"billsub":786714},"2026-05":{"orders":1942,"revenue":393091,"billsub":466080}},"item_count":59,"items":[{"name":"Sweet Lassi","cat":"Beverages","price":90,"orders":800,"rating":4.3,"fc":18},{"name":"Jeera Rice","cat":"Rice & Biryani","price":180,"orders":714,"rating":4.12,"fc":22},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":200,"orders":669,"rating":3.64,"fc":24},{"name":"Hara Bhara Kabab","cat":"Starters","price":220,"orders":650,"rating":3.6,"fc":28},{"name":"Paneer Tikka","cat":"Starters","price":380,"orders":618,"rating":3.81,"fc":34},{"name":"Veg Spring Roll","cat":"Starters","price":210,"orders":607,"rating":3.81,"fc":24},{"name":"Medu Vada","cat":"South Indian","price":110,"orders":597,"rating":4.5,"fc":22},{"name":"Triple Schezwan Rice","cat":"Asian Mains","price":320,"orders":578,"rating":3.71,"fc":24},{"name":"Idli Sambar","cat":"South Indian","price":110,"orders":494,"rating":4.04,"fc":20},{"name":"Plain Dosa","cat":"South Indian","price":100,"orders":411,"rating":3.7,"fc":22},{"name":"Uttapam","cat":"South Indian","price":160,"orders":326,"rating":3.63,"fc":26},{"name":"Veg Manchurian","cat":"Asian Mains","price":210,"orders":311,"rating":3.65,"fc":22},{"name":"Veg Hakka Noodles","cat":"Asian Mains","price":250,"orders":283,"rating":4.19,"fc":20},{"name":"Mushroom Tikka","cat":"Starters","price":260,"orders":278,"rating":3.93,"fc":30},{"name":"Crispy Corn","cat":"Starters","price":230,"orders":218,"rating":3.71,"fc":22},{"name":"Veg Seekh","cat":"Starters","price":300,"orders":188,"rating":4.59,"fc":30},{"name":"Mango Lassi","cat":"Beverages","price":140,"orders":178,"rating":3.99,"fc":22},{"name":"Rava Dosa","cat":"South Indian","price":150,"orders":165,"rating":4.55,"fc":27},{"name":"Masala Chaas","cat":"Beverages","price":70,"orders":158,"rating":3.7,"fc":14},{"name":"Steamed Rice","cat":"Rice & Biryani","price":120,"orders":151,"rating":4.05,"fc":18},{"name":"Masala Dosa","cat":"South Indian","price":150,"orders":102,"rating":4.0,"fc":25}],"mBillSub":602972,"mCV":498849,"mDiscount":77867,"mOrders":2030},{"id":"MOCK1002","name":"Amber Bistro","subzone":"Maple Heights","cuisine":"fast food","monthly":{"2026-02":{"orders":1792,"revenue":411105,"billsub":507136},"2026-03":{"orders":1847,"revenue":529669,"billsub":670461},"2026-04":{"orders":1867,"revenue":566557,"billsub":705726},"2026-05":{"orders":2065,"revenue":457855,"billsub":578200}},"item_count":62,"items":[{"name":"Brownie","cat":"Desserts","price":140,"orders":799,"rating":4.04,"fc":32},{"name":"Crispy Corn","cat":"Starters","price":200,"orders":700,"rating":3.9,"fc":22},{"name":"Paneer Tikka","cat":"Starters","price":320,"orders":637,"rating":4.17,"fc":34},{"name":"Hara Bhara Kabab","cat":"Starters","price":260,"orders":616,"rating":4.6,"fc":28},{"name":"Rasmalai","cat":"Desserts","price":150,"orders":614,"rating":3.61,"fc":30},{"name":"Kadhai Paneer","cat":"Main Course","price":300,"orders":605,"rating":3.75,"fc":33},{"name":"Jeera Rice","cat":"Rice & Biryani","price":160,"orders":596,"rating":4.3,"fc":22},{"name":"Curd Rice","cat":"Rice & Biryani","price":170,"orders":418,"rating":3.88,"fc":20},{"name":"Dal Khichdi","cat":"Rice & Biryani","price":200,"orders":397,"rating":4.48,"fc":24},{"name":"Paneer Butter Masala","cat":"Main Course","price":270,"orders":355,"rating":3.96,"fc":32},{"name":"Ice Cream","cat":"Desserts","price":100,"orders":323,"rating":3.64,"fc":24},{"name":"Veg Seekh","cat":"Starters","price":250,"orders":307,"rating":3.8,"fc":30},{"name":"Veg Spring Roll","cat":"Starters","price":200,"orders":284,"rating":3.87,"fc":24},{"name":"Garlic Naan","cat":"Breads","price":90,"orders":199,"rating":3.87,"fc":18},{"name":"Gulab Jamun","cat":"Desserts","price":120,"orders":173,"rating":4.24,"fc":28},{"name":"Butter Roti","cat":"Breads","price":40,"orders":158,"rating":4.49,"fc":13},{"name":"Cheese Naan","cat":"Breads","price":130,"orders":146,"rating":4.54,"fc":22},{"name":"Mushroom Tikka","cat":"Starters","price":270,"orders":114,"rating":3.61,"fc":30},{"name":"Gajar Halwa","cat":"Desserts","price":150,"orders":92,"rating":3.69,"fc":30},{"name":"Chana Masala","cat":"Main Course","price":250,"orders":80,"rating":4.5,"fc":24}],"mBillSub":615381,"mCV":491296,"mDiscount":98226,"mOrders":1893}],"benchmarks":{"paneertikka":[332.2,282.4,392.0,18],"hyderabadibiryani":[323.8,275.2,382.0,16],"butterroti":[43.0,36.5,50.7,20],"tandooriroti":[36.2,30.8,42.8,16],"butternaan":[61.2,52.1,72.3,16],"paneerchillidry":[285.6,242.7,337.0,18],"kadhaipaneer":[311.0,264.3,367.0,20],"veghakkanoodles":[237.0,201.4,279.7,20],"brownie":[155.7,132.4,183.7,14],"vegkolhapuri":[296.0,251.6,349.3,10],"vegpulao":[211.7,179.9,249.8,12],"jeerarice":[173.3,147.3,204.5,18],"vegchowmein":[211.1,179.4,249.1,18],"vegspringroll":[190.0,161.5,224.2,26],"chillipaneer":[280.0,238.0,330.4,18],"vegbiryani":[240.0,204.0,283.2,20],"gulabjamun":[125.0,106.2,147.5,20],"mushroomtikka":[282.5,240.1,333.3,24],"paneerbuttermasala":[308.6,262.3,364.1,14],"crispycorn":[221.7,188.4,261.6,24],"cheesenaan":[111.1,94.4,131.1,18],"palakpaneer":[302.5,257.1,356.9,16],"garlicnaan":[83.0,70.5,97.9,20],"vegmanchurian":[240.0,204.0,283.2,18],"harabharakabab":[240.0,204.0,283.2,24],"chanamasala":[222.5,189.1,262.6,16],"coldcoffee":[128.9,109.6,152.1,18],"masalachaas":[64.2,54.5,75.7,24],"vegseekh":[262.9,223.4,310.2,14],"sweetlassi":[92.2,78.4,108.8,18],"chillimushroom":[253.3,215.3,298.9,12],"schezwanfriedrice":[234.4,199.3,276.6,18],"icecream":[99.1,84.2,116.9,22],"curdrice":[161.8,137.5,190.9,22],"gajarhalwa":[163.3,138.8,192.7,24],"dalmakhani":[246.7,209.7,291.1,12],"malaikofta":[304.4,258.8,359.2,18],"paneerbiryani":[305.0,259.2,359.9,16],"dalkhichdi":[219.1,186.2,258.5,22],"steamedrice":[120.0,102.0,141.6,18],"ravadosa":[152.7,129.8,180.2,22],"meduvada":[108.0,91.8,127.4,20],"mixveg":[231.7,196.9,273.4,12],"plaindosa":[116.0,98.6,136.9,20],"idlisambar":[111.0,94.3,131.0,20],"tripleschezwanrice":[294.0,249.9,346.9,20],"uttapam":[150.0,127.5,177.0,22],"masaladosa":[147.5,125.4,174.0,24],"freshlimesoda":[75.7,64.4,89.3,14],"rasmalai":[150.8,128.2,178.0,24],"mangolassi":[120.0,102.0,141.6,18],"lachhaparatha":[76.2,64.8,90.0,16],"shahipaneer":[303.3,257.8,357.9,12]},"mock":true};

// ============ HELPERS ============
const inr = (n) => "₹" + Math.round(n || 0).toLocaleString("en-IN");
const inrCompact = (n) => {
  if (n == null || isNaN(n)) return "-";
  const a = Math.abs(n);
  if (a >= 1e7) return "₹" + (n / 1e7).toFixed(2) + "Cr";
  if (a >= 1e5) return "₹" + (n / 1e5).toFixed(2) + "L";
  if (a >= 1e3) return "₹" + (n / 1e3).toFixed(1) + "K";
  return "₹" + Math.round(n);
};
const num = (n) => (n || 0).toLocaleString("en-IN");
const normName = (s) => (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
const NUM_MONTHS = 4;

const lookupBenchmark = (name) => {
  const k = normName(name);
  const b = DATA.benchmarks[k];
  if (!b) return null;
  return { median: b[0], p25: b[1], p75: b[2], n: b[3] };
};

// ===== Promotability: items that CAN be pushed via ads/discounts/combos =====
const NON_PROMOTABLE_CATS = ["Breads", "Sides & Add-Ons", "Salads & Sides"];
const NON_PROMOTABLE_KEYWORDS = [
  "plain rice", "steamed rice", "jeera rice", "curd rice", "roti", "naan",
  "paratha", "papad", "chapati", "kulcha", "bread", "rumali", "extra ",
  "add on", "add-on", "water", "raita",
];
function isPromotable(item) {
  if (NON_PROMOTABLE_CATS.includes(item.cat || item.category)) return false;
  const n = (item.name || "").toLowerCase();
  return !NON_PROMOTABLE_KEYWORDS.some((kw) => n.includes(kw));
}

// ===== CODB / payout math =====
function computeItemMargin(menuPrice, foodCostPct, codbPct) {
  const foodCost = (menuPrice * foodCostPct) / 100;
  const grossPayout = menuPrice * (1 - codbPct / 100);
  const net = grossPayout - foodCost;
  return {
    foodCost, grossPayout, net,
    marginPct: menuPrice > 0 ? (net / menuPrice) * 100 : 0,
  };
}

const DEFAULT_BREAKDOWN = { commission: 20, ads: 8, discount: 12, pg: 2, gst: 18 };
function computeCODB(b) {
  return b.commission + b.ads + b.discount + b.pg + ((b.commission + b.ads + b.pg) * b.gst) / 100;
}

const DEFAULT_FC_BY_CAT = {
  "Rice & Biryani": 30, "Main Course": 32, "Starters": 30, "Breads": 16,
  "South Indian": 25, "Desserts": 28, "Beverages": 18, "Burgers & Sandwiches": 35,
  "Rolls & Wraps": 32, "Pizza & Pasta": 35, "Momos & Dimsum": 30, "Noodles": 24,
  "Asian Mains": 23, "Street Food": 25, "Combos & Thalis": 35, "Soups": 22,
  "Salads & Sides": 25, "Sides & Add-Ons": 18, "Breakfast": 25, "Snacks": 25, "Other": 30,
};

// ============ MENU PARSING (CSV / XLSX / DOCX / PDF / IMAGE-OCR / TEXT) ============
function normHeader(h) { return String(h || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
function findField(row, candidates) {
  const map = {};
  for (const k in row) map[normHeader(k)] = row[k];
  for (const c of candidates) {
    const k = normHeader(c);
    if (map[k] !== undefined && map[k] !== null && String(map[k]).trim() !== "") return map[k];
  }
  return null;
}
function normalizeCat(c) {
  c = (c || "").toLowerCase();
  if (/biryani|pulao|fried rice|^rice|jeera|khichdi/.test(c)) return "Rice & Biryani";
  if (/starter|tikka|kabab|kebab|crispy|spring roll|65/.test(c)) return "Starters";
  if (/manchurian|chilli|hakka|noodle|schezwan|chowmein|chinese/.test(c)) return "Asian Mains";
  if (/dessert|sweet|gulab|halwa|ice cream|rasmalai|brownie/.test(c)) return "Desserts";
  if (/beverage|drink|lassi|chaas|soda|coffee|tea|juice/.test(c)) return "Beverages";
  if (/roti|naan|paratha|bread|kulcha/.test(c)) return "Breads";
  if (/dosa|idli|vada|uttapam|south/.test(c)) return "South Indian";
  if (/paneer|dal|kofta|masala|gravy|main|kadhai|veg /.test(c)) return "Main Course";
  return "Other";
}
function extractItemsFromRows(rows) {
  const items = [];
  for (const r of rows) {
    if (!r || typeof r !== "object") continue;
    const name = findField(r, ["item name", "itemname", "name", "item", "dish", "dish name", "product", "menu item", "items"]);
    const catRaw = findField(r, ["category", "menu category", "section", "type", "menu section"]);
    const priceRaw = findField(r, ["selling price", "price", "menu price", "rate", "amount", "mrp", "cost", "sp", "avg unit cost", "avg_unit_cost", "online price", "offline price", "dine in price"]);
    const price = parseFloat(String(priceRaw).replace(/[^0-9.]/g, ""));
    if (!name || isNaN(price) || price <= 0) continue;
    const cat = catRaw ? normalizeCat(catRaw) : normalizeCat(name);
    items.push({ name: String(name).trim(), category: cat, price, foodCostPct: DEFAULT_FC_BY_CAT[cat] || 30 });
  }
  return items;
}
function parseMenuText(text) {
  const lines = text.split(/\n/).map((l) => l.trim()).filter(Boolean);
  const items = [];
  let currentCat = "Other";
  // Match a trailing price: "Paneer Tikka .... 320" or "Paneer Tikka - ₹320" or "320/-"
  const priceRe = /(?:₹|rs\.?|inr)?\s*(\d{2,4}(?:\.\d{1,2})?)\s*\/?-?\s*$/i;
  const headerRe = /^[A-Z][A-Za-z\s&'-]{2,28}$/;
  for (let line of lines) {
    line = line.replace(/\.{2,}/g, " ").trim();
    if (line.length < 3) continue;
    const m = line.match(priceRe);
    if (!m) {
      // possible category header (short, no digits)
      if (line.length <= 28 && headerRe.test(line) && !/\d/.test(line)) {
        currentCat = normalizeCat(line);
      }
      continue;
    }
    const price = parseFloat(m[1]);
    if (price < 20 || price > 5000) continue;
    let name = line.slice(0, m.index).replace(/[.\--–:]+$/, "").trim();
    if (name.length < 2) continue;
    const cat = normalizeCat(name) !== "Other" ? normalizeCat(name) : currentCat;
    items.push({ name, category: cat, price, foodCostPct: DEFAULT_FC_BY_CAT[cat] || 30 });
  }
  return items;
}

async function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: "greedy",
      transformHeader: (h) => String(h || "").trim(),
      complete: (res) => resolve(extractItemsFromRows(res.data)),
      error: reject,
    });
  });
}
async function parseXLSX(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  let best = [];
  for (const sn of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: "" });
    const items = extractItemsFromRows(rows);
    if (items.length > best.length) best = items;
  }
  return best;
}
async function parseDOCX(file) {
  const buf = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: buf });
  return parseMenuText(result.value);
}
async function loadScript(src) {
  return new Promise((resolve, reject) => {
    if ([...document.scripts].some((s) => s.src === src)) return resolve();
    const el = document.createElement("script");
    el.src = src; el.onload = resolve; el.onerror = () => reject(new Error("load failed: " + src));
    document.head.appendChild(el);
  });
}
async function parsePDF(file) {
  if (!window.pdfjsLib) {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
    if (window.pdfjsLib) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    }
  }
  const lib = window.pdfjsLib;
  if (!lib) throw new Error("PDF reader unavailable");
  const buf = await file.arrayBuffer();
  const pdf = await lib.getDocument({ data: buf }).promise;
  let text = "";
  let hasText = false;
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items.map((it) => it.str).join("\n");
    if (pageText.trim().length > 20) hasText = true;
    text += pageText + "\n";
  }
  // If PDF is scanned (no text layer), OCR the rendered pages
  if (!hasText) {
    let ocrText = "";
    for (let i = 1; i <= Math.min(pdf.numPages, 5); i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width; canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      ocrText += await ocrImageData(canvas) + "\n";
    }
    return parseMenuText(ocrText);
  }
  return parseMenuText(text);
}
async function ocrImageData(canvasOrUrl) {
  if (!window.Tesseract) {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/5.0.4/tesseract.min.js");
  }
  const T = window.Tesseract;
  if (!T) throw new Error("OCR engine unavailable");
  const { data } = await T.recognize(canvasOrUrl, "eng");
  return data.text || "";
}
async function parseImage(file) {
  const url = URL.createObjectURL(file);
  try {
    const text = await ocrImageData(url);
    return parseMenuText(text);
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

// Master dispatcher
async function parseMenuFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  if (ext === "csv") return parseCSV(file);
  if (ext === "xlsx" || ext === "xls") return parseXLSX(file);
  if (ext === "docx") return parseDOCX(file);
  if (ext === "pdf") return parsePDF(file);
  if (["png", "jpg", "jpeg", "webp", "bmp", "gif", "heic"].includes(ext)) return parseImage(file);
  // last resort: try reading as text
  const text = await file.text();
  return parseMenuText(text);
}

// Fuzzy match an uploaded menu to a restaurant in the DB (by item-name overlap)
function matchRestaurant(uploadedItems) {
  if (!uploadedItems || uploadedItems.length === 0) return null;
  const upSet = new Set(uploadedItems.map((i) => normName(i.name)));
  let best = null, bestScore = 0;
  for (const r of DATA.restaurants) {
    let hits = 0;
    for (const it of r.items) if (upSet.has(normName(it.name))) hits++;
    const score = hits / Math.max(1, r.items.length);
    if (score > bestScore) { bestScore = score; best = r; }
  }
  return bestScore >= 0.15 ? { restaurant: best, score: bestScore } : null;
}

// ============ MAIN APP ============
export default function App() {
  const [mode, setMode] = useState("partner"); // 'partner' | 'admin'
  const [resId, setResId] = useState(DATA.restaurants[0].id);
  const [breakdown, setBreakdown] = useState(DEFAULT_BREAKDOWN);
  const [codb, setCodb] = useState(Math.round(computeCODB(DEFAULT_BREAKDOWN)));
  const [advanced, setAdvanced] = useState(false);
  const [targetMargin, setTargetMargin] = useState(20);
  const [theme, setTheme] = useState("light");
  const [uploadOpen, setUploadOpen] = useState(false);
  const [menuOverride, setMenuOverride] = useState(null);
  const [matchInfo, setMatchInfo] = useState(null);
  const [inflationData, setInflationData] = useState(null); // {items:[{name, offline, online, infl}], blended}

  const dark = theme === "dark";
  const restaurant = useMemo(() => DATA.restaurants.find((r) => r.id === resId), [resId]);

  useEffect(() => { setCodb(+computeCODB(breakdown).toFixed(1)); }, [breakdown]);
  useEffect(() => { setMenuOverride(null); setMatchInfo(null); setInflationData(null); }, [resId]);

  const menu = useMemo(() => {
    if (menuOverride) return menuOverride;
    if (!restaurant) return [];
    return restaurant.items.map((it) => ({
      name: it.name, category: it.cat, price: it.price,
      foodCostPct: it.fc ?? DEFAULT_FC_BY_CAT[it.cat] ?? 30,
      monthlyOrders: Math.round(it.orders / NUM_MONTHS),
      rating: it.rating,
    }));
  }, [restaurant, menuOverride]);

  // Per-item profitability
  const rows = useMemo(() => {
    return menu.map((item, idx) => {
      const m = computeItemMargin(item.price, item.foodCostPct, codb);
      const bench = lookupBenchmark(item.name);
      const monthlyProfit = m.net * (item.monthlyOrders || 0);
      const monthlyRev = item.price * (item.monthlyOrders || 0);
      return { ...item, idx, ...m, monthlyProfit, monthlyRev, bench, promotable: isPromotable(item) };
    });
  }, [menu, codb]);

  // Restaurant rollup (weighted)
  const rollup = useMemo(() => {
    const rev = rows.reduce((s, r) => s + r.monthlyRev, 0);
    const profit = rows.reduce((s, r) => s + r.monthlyProfit, 0);
    return { rev, profit, blendedMargin: rev ? (profit / rev) * 100 : 0 };
  }, [rows]);

  // Cuisine benchmark margin - average blended margin of DB restaurants in the same cuisine
  const cuisineBenchmark = useMemo(() => {
    if (!restaurant) return null;
    const peers = DATA.restaurants.filter((r) => r.cuisine === restaurant.cuisine);
    if (peers.length === 0) return null;
    const margins = peers.map((r) => {
      let rev = 0, profit = 0;
      r.items.forEach((it) => {
        const mo = it.orders / NUM_MONTHS;
        const m = computeItemMargin(it.price, it.fc ?? DEFAULT_FC_BY_CAT[it.cat] ?? 30, codb);
        rev += it.price * mo; profit += m.net * mo;
      });
      return rev ? (profit / rev) * 100 : 0;
    }).filter((m) => m > 0);
    if (!margins.length) return null;
    const avg = margins.reduce((s, m) => s + m, 0) / margins.length;
    return { avg, cuisine: restaurant.cuisine, peers: peers.length };
  }, [restaurant, codb]);

  // ===== TOP PROFIT OPPORTUNITY - grouped by CATEGORY =====
  // Each promotable category gets an opportunity score; expandable to see its items.
  const opportunityCategories = useMemo(() => {
    const promotable = rows.filter((r) => r.promotable);
    if (promotable.length === 0) return [];
    const maxOrders = Math.max(...rows.map((r) => r.monthlyOrders), 1);

    // group by category
    const groups = {};
    promotable.forEach((r) => { (groups[r.category] = groups[r.category] || []).push(r); });

    const cats = Object.entries(groups).map(([cat, items]) => {
      const avgMargin = items.reduce((s, r) => s + r.marginPct, 0) / items.length;
      const totalOrders = items.reduce((s, r) => s + r.monthlyOrders, 0);
      const totalProfit = items.reduce((s, r) => s + r.monthlyProfit, 0);
      // items above target margin that sell relatively little = the real opportunity
      const oppItems = items
        .filter((r) => r.marginPct >= targetMargin)
        .map((r) => ({ ...r, volumeRoom: 1 - r.monthlyOrders / maxOrders }))
        .sort((a, b) => (b.marginPct * 0.6 + b.volumeRoom * 40) - (a.marginPct * 0.6 + a.volumeRoom * 40));
      // category opportunity score: high margin + has headroom + meaningful profit-per-order
      const marginScore = Math.min(1, avgMargin / 35);
      const headroom = 1 - totalOrders / (maxOrders * items.length);
      const score = Math.round((marginScore * 0.6 + headroom * 0.4) * 100);
      return {
        cat, items: items.sort((a, b) => b.marginPct - a.marginPct),
        oppItems, avgMargin, totalOrders, totalProfit, score,
        hasOpportunity: oppItems.length > 0,
      };
    });

    return cats
      .filter((c) => c.hasOpportunity)
      .sort((a, b) => b.score - a.score);
  }, [rows, targetMargin]);

  // Flat list of opportunity items (for the promotion simulator default + insights)
  const opportunities = useMemo(() => {
    const flat = [];
    opportunityCategories.forEach((c) => c.oppItems.forEach((it) => flat.push(it)));
    return flat.sort((a, b) => b.marginPct - a.marginPct).slice(0, 25);
  }, [opportunityCategories]);

  // ===== MARGIN MIX (category-level) =====
  const categoryMix = useMemo(() => {
    const cats = {};
    rows.forEach((r) => {
      if (!cats[r.category]) cats[r.category] = { cat: r.category, rev: 0, profit: 0, orders: 0, items: 0, promotable: 0 };
      const c = cats[r.category];
      c.rev += r.monthlyRev; c.profit += r.monthlyProfit; c.orders += r.monthlyOrders; c.items++;
      if (r.promotable) c.promotable++;
    });
    const totalRev = Object.values(cats).reduce((s, c) => s + c.rev, 0);
    const totalProfit = Object.values(cats).reduce((s, c) => s + c.profit, 0);
    return Object.values(cats).map((c) => ({
      ...c,
      margin: c.rev ? (c.profit / c.rev) * 100 : 0,
      revShare: totalRev ? (c.rev / totalRev) * 100 : 0,
      profitShare: totalProfit ? (c.profit / totalProfit) * 100 : 0,
    })).sort((a, b) => b.profitShare - a.profitShare);
  }, [rows]);

  // ===== INSIGHTS (mix/promotion only, no pricing) =====
  const insights = useMemo(() => {
    const out = [];
    // 1. High-margin category under-indexed in sales
    const underIndexed = categoryMix.find((c) => c.profitShare > c.revShare + 5 && c.margin > targetMargin && c.promotable > 0);
    if (underIndexed) {
      out.push({
        tone: "opportunity",
        title: `${underIndexed.cat} punches above its weight`,
        body: `${underIndexed.cat} brings ${underIndexed.revShare.toFixed(0)}% of your sales but ${underIndexed.profitShare.toFixed(0)}% of your profit. Promoting it more - combos, banners, recommendations - can lift your blended margin without touching prices.`,
      });
    }
    // 2. Volume category with thin margin (don't push, just note)
    const volumeThin = categoryMix.find((c) => c.revShare > 18 && c.margin < targetMargin - 5);
    if (volumeThin) {
      out.push({
        tone: "info",
        title: `${volumeThin.cat} drives volume more than profit`,
        body: `${volumeThin.cat} is ${volumeThin.revShare.toFixed(0)}% of sales but only ${volumeThin.profitShare.toFixed(0)}% of profit. Keep it selling, but spend your ad and discount budget on higher-margin items instead.`,
      });
    }
    // 3. Best single opportunity item
    if (opportunities.length > 0) {
      const top = opportunities[0];
      out.push({
        tone: "opportunity",
        title: `${top.name} is a strong item to promote`,
        body: `It earns a healthy ${top.marginPct.toFixed(0)}% margin but sells only ${num(top.monthlyOrders)}/month. More visibility here improves profitability without raising any prices.`,
      });
    }
    // 4. Two-category comparison
    const promotableCats = categoryMix.filter((c) => c.promotable > 0 && c.orders > 0);
    if (promotableCats.length >= 2) {
      const best = [...promotableCats].sort((a, b) => b.margin - a.margin)[0];
      const worst = [...promotableCats].sort((a, b) => a.margin - b.margin)[0];
      if (best.margin - worst.margin > 6) {
        out.push({
          tone: "info",
          title: `${best.cat} is more profitable than ${worst.cat}`,
          body: `${best.cat} runs at ${best.margin.toFixed(0)}% margin vs ${worst.cat} at ${worst.margin.toFixed(0)}%. Shifting promotion budget toward ${best.cat} improves your overall mix.`,
        });
      }
    }
    // 5. Philosophy
    out.push({
      tone: "info",
      title: `Better mix beats more orders`,
      body: `You don't need to sell more overall - selling more of your high-margin items lifts blended profitability on its own. Use the simulator below to see the effect.`,
    });
    return out.slice(0, 5);
  }, [categoryMix, opportunities, targetMargin]);

  // Theme tokens
  const t = useMemo(() => ({
    bg: dark ? "bg-zinc-950" : "bg-stone-50",
    card: dark ? "bg-zinc-900 border-zinc-800" : "bg-white border-zinc-200",
    text: dark ? "text-zinc-50" : "text-zinc-950",
    muted: dark ? "text-zinc-400" : "text-zinc-500",
    border: dark ? "border-zinc-800" : "border-zinc-200",
    softBg: dark ? "bg-zinc-900/50" : "bg-zinc-100/70",
    pos: dark ? "text-lime-400" : "text-emerald-700",
    neg: dark ? "text-rose-400" : "text-rose-700",
    warn: dark ? "text-amber-300" : "text-amber-700",
    posBg: dark ? "bg-emerald-950/40 border-emerald-800/40" : "bg-emerald-50 border-emerald-200",
    negBg: dark ? "bg-rose-950/40 border-rose-800/40" : "bg-rose-50 border-rose-200",
    warnBg: dark ? "bg-amber-950/40 border-amber-800/40" : "bg-amber-50 border-amber-200",
    infoBg: dark ? "bg-blue-950/40 border-blue-800/40" : "bg-blue-50 border-blue-200",
    accent: dark ? "bg-lime-400 text-zinc-950" : "bg-zinc-950 text-white",
  }), [dark]);

  const marginTone = (m) => (m < 8 ? "neg" : m < targetMargin ? "warn" : "pos");

  // Handle menu import + auto-match
  const handleImport = ({ online, offline }) => {
    const primary = online && online.length ? online : offline;
    if (primary && primary.length) {
      setMenuOverride(primary.map((i) => ({
        name: i.name, category: i.category, price: i.price,
        foodCostPct: i.foodCostPct ?? DEFAULT_FC_BY_CAT[i.category] ?? 30,
        monthlyOrders: 0, rating: 0,
      })));
      const match = matchRestaurant(primary);
      setMatchInfo(match);
      if (match) setResId(match.restaurant.id);
    }
    // Price inflation comparison if both provided
    if (online && online.length && offline && offline.length) {
      const offMap = {};
      offline.forEach((i) => { offMap[normName(i.name)] = i.price; });
      const items = [];
      let onSum = 0, offSum = 0;
      online.forEach((i) => {
        const off = offMap[normName(i.name)];
        if (off && off > 0) {
          const infl = ((i.price - off) / off) * 100;
          items.push({ name: i.name, category: i.category, offline: off, online: i.price, infl });
          onSum += i.price; offSum += off;
        }
      });
      const blended = offSum ? ((onSum - offSum) / offSum) * 100 : 0;
      setInflationData({ items: items.sort((a, b) => b.infl - a.infl), blended, matched: items.length });
    } else {
      setInflationData(null);
    }
    setUploadOpen(false);
  };

  if (mode === "admin") {
    return <AdminPanel t={t} dark={dark} onExit={() => setMode("partner")} onThemeToggle={() => setTheme(dark ? "light" : "dark")} />;
  }

  return (
    <div className={`min-h-screen ${t.bg} ${t.text} antialiased`}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        * { font-family: 'Inter', system-ui, sans-serif; }
        .num { font-variant-numeric: tabular-nums; }
        .tight { letter-spacing: -0.03em; }
        .ultra-tight { letter-spacing: -0.05em; }
        input[type="range"] { -webkit-appearance: none; height: 8px; border-radius: 999px; outline: none; }
        input[type="range"]::-webkit-slider-thumb { -webkit-appearance: none; width: 24px; height: 24px; border-radius: 50%; background: ${dark ? "#a3e635" : "#0a0a0a"}; cursor: pointer; border: 3px solid ${dark ? "#0a0a0a" : "#fff"}; box-shadow: 0 2px 6px rgba(0,0,0,0.15); }
        input[type="range"]::-moz-range-thumb { width: 24px; height: 24px; border-radius: 50%; background: ${dark ? "#a3e635" : "#0a0a0a"}; cursor: pointer; border: 3px solid ${dark ? "#0a0a0a" : "#fff"}; }
      `}</style>

      {/* MOCK DATA DISCLAIMER */}
      <MockBanner />

      {/* HEADER */}
      <header className={`border-b ${t.border} sticky top-0 z-30 ${t.bg}`}>
        <div className="max-w-5xl mx-auto px-4 md:px-6 h-14 flex items-center gap-3">
          <div className="flex items-center gap-2 shrink-0">
            <div className={`w-7 h-7 ${t.accent} rounded-md flex items-center justify-center font-black text-sm ultra-tight`}>M</div>
            <div className="hidden sm:block">
              <div className="font-black text-base tight leading-none">Menu Mix Assistant</div>
              <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-bold mt-0.5`}>Grow profit through better mix</div>
            </div>
          </div>
          <select value={resId} onChange={(e) => setResId(e.target.value)}
            className={`ml-auto md:ml-4 ${t.softBg} ${t.text} border ${t.border} rounded-lg px-3 py-1.5 text-sm font-bold max-w-[200px] sm:max-w-none truncate`}>
            {DATA.restaurants.map((r) => <option key={r.id} value={r.id}>{r.name} - {r.subzone}</option>)}
          </select>
          <button onClick={() => setMode("admin")}
            className={`hidden sm:block text-xs font-bold px-3 py-1.5 rounded-md border ${t.border} hover:${t.softBg}`}>
            KAM Admin
          </button>
          <button onClick={() => setTheme(dark ? "light" : "dark")}
            className={`text-xs font-bold px-2 py-1.5 rounded-md border ${t.border} hover:${t.softBg}`}>
            {dark ? "☀" : "☾"}
          </button>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 md:px-6 py-6 md:py-10 space-y-6 md:space-y-8">
        {/* HERO */}
        <section>
          <div className={`text-[11px] uppercase tracking-widest ${t.muted} font-black mb-2`}>You're looking at</div>
          <h1 className="text-2xl md:text-4xl font-black ultra-tight leading-tight">{restaurant.name}</h1>
          <div className={`text-sm ${t.muted} font-semibold mt-1`}>
            {restaurant.subzone} · <span className="capitalize">{restaurant.cuisine}</span> · {menuOverride ? "Your uploaded menu" : `${menu.length} dishes`}
            {matchInfo && <span className={`ml-1 ${t.pos} font-bold`}>· matched to {matchInfo.restaurant.name} ({Math.round(matchInfo.score * 100)}% items)</span>}
          </div>
        </section>

        {/* P&L SNAPSHOT */}
        <PLSnapshot rollup={rollup} restaurant={restaurant} codb={codb} t={t} />

        {/* UPLOAD */}
        <UploadStep t={t} onOpen={() => setUploadOpen(true)} hasOverride={!!menuOverride} onClear={() => { setMenuOverride(null); setInflationData(null); setMatchInfo(null); }} />

        {/* PRICE INFLATION (if dual menu uploaded) */}
        {inflationData && <InflationView data={inflationData} t={t} dark={dark} />}

        {/* CODB */}
        <CodbStep codb={codb} setCodb={setCodb} advanced={advanced} setAdvanced={setAdvanced} breakdown={breakdown} setBreakdown={setBreakdown} t={t} />

        {/* TARGET MARGIN */}
        <TargetMarginStep targetMargin={targetMargin} setTargetMargin={setTargetMargin} cuisineBenchmark={cuisineBenchmark} blendedMargin={rollup.blendedMargin} t={t} dark={dark} />

        {/* TOP PROFIT OPPORTUNITY CATEGORIES */}
        <OpportunityCategories categories={opportunityCategories} t={t} dark={dark} targetMargin={targetMargin} />

        {/* PROMOTION IMPACT SIMULATOR */}
        <PromotionSimulator opportunities={opportunities} rows={rows} rollup={rollup} codb={codb} t={t} dark={dark} />

        {/* MARGIN MIX */}
        <MarginMix categoryMix={categoryMix} t={t} dark={dark} targetMargin={targetMargin} />

        {/* INSIGHTS */}
        <InsightsView insights={insights} t={t} />

        <div className={`text-xs ${t.muted} font-medium text-center pt-4 pb-2`}>
          <strong className={t.text}>All figures are mock / representational only.</strong> This is not actual restaurant data.
        </div>
      </main>

      {uploadOpen && (
        <UploadModal onClose={() => setUploadOpen(false)} onImport={handleImport} t={t} dark={dark} />
      )}
    </div>
  );
}

// ============ MOCK DATA BANNER ============
function MockBanner() {
  return (
    <div className="w-full bg-amber-400 text-zinc-950 text-center py-1.5 px-3 text-[10px] sm:text-xs font-black uppercase tracking-wider sticky top-0 z-40">
      ⚠ ALL DATA SHOWN HERE IS MOCK / REPRESENTATIONAL DATA ONLY · NOT ACTUAL RESTAURANT DATA
    </div>
  );
}

// ============ P&L SNAPSHOT ============
function PLSnapshot({ rollup, restaurant, codb, t }) {
  const monthly = restaurant.monthly ? Object.entries(restaurant.monthly).sort(([a], [b]) => a.localeCompare(b)) : [];
  const billed = restaurant.mBillSub || 0;
  const realProfit = restaurant.mCV ? restaurant.mCV * (1 - codb / 100) * 0.68 : rollup.profit; // approx after food cost
  const marginTone = rollup.blendedMargin < 8 ? t.neg : rollup.blendedMargin < 15 ? t.warn : t.pos;
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-3`}>Your business right now</div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-4 md:gap-6">
        <div>
          <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1.5 flex items-center gap-1`}>
            Monthly billed subtotal <InfoTip t={t} text="Your total monthly online sales at menu prices, before discounts and taxes." />
          </div>
          <div className="text-2xl md:text-3xl font-black num ultra-tight">{inrCompact(billed)}</div>
        </div>
        <div>
          <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1.5`}>Blended margin</div>
          <div className={`text-2xl md:text-3xl font-black num ultra-tight ${marginTone}`}>{rollup.blendedMargin.toFixed(1)}%</div>
          <div className={`text-[11px] ${t.muted} font-bold mt-0.5`}>across your menu</div>
        </div>
        <div className="col-span-2 md:col-span-1">
          <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1.5`}>Monthly orders</div>
          <div className="text-2xl md:text-3xl font-black num ultra-tight">{num(restaurant.mOrders || 0)}</div>
          <div className={`text-[11px] ${t.muted} font-bold mt-0.5`}>on the platform</div>
        </div>
      </div>
    </section>
  );
}

// ============ UPLOAD STEP ============
function UploadStep({ t, onOpen, hasOverride, onClear }) {
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <div className="flex items-baseline gap-3 mb-1">
        <div className={`text-xs font-black ${t.muted}`}>YOUR MENU</div>
        <h2 className="text-lg md:text-xl font-black ultra-tight">Upload your menu</h2>
      </div>
      <p className={`text-sm ${t.muted} font-medium mb-4`}>
        Upload a photo, PDF, Word, Excel, or CSV of your menu - we'll read it automatically. Add both your dine-in and online menus to see exactly how much your online prices are marked up.
      </p>
      <div className="flex flex-wrap gap-2">
        <button onClick={onOpen} className={`text-sm font-bold px-4 py-2.5 rounded-lg ${t.accent}`}>
          Upload Menu
        </button>
        {hasOverride && (
          <button onClick={onClear} className={`text-xs font-bold px-3 py-2 rounded-lg border ${t.border} ${t.muted} hover:${t.softBg}`}>
            Use sample data instead
          </button>
        )}
      </div>
    </section>
  );
}

// ============ PRICE INFLATION VIEW ============
function InflationView({ data, t, dark }) {
  const tone = data.blended > 15 ? t.warn : data.blended > 0 ? t.pos : t.muted;
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <h2 className="text-lg md:text-xl font-black ultra-tight mb-1">Online vs dine-in pricing</h2>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>
        How much higher your online prices are versus your dine-in menu, matched on {data.matched} item{data.matched > 1 ? "s" : ""}.
      </p>
      <div className={`rounded-xl border p-4 mb-5 ${data.blended > 15 ? t.warnBg : t.posBg}`}>
        <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1`}>Overall menu mark-up</div>
        <div className={`text-4xl font-black num ultra-tight ${tone}`}>{data.blended >= 0 ? "+" : ""}{data.blended.toFixed(1)}%</div>
        <div className={`text-xs font-bold ${t.muted} mt-1`}>
          Your online menu is on average {Math.abs(data.blended).toFixed(0)}% {data.blended >= 0 ? "higher" : "lower"} than dine-in
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className={`border-b ${t.border}`}>
              {["Item", "Dine-in", "Online", "Mark-up"].map((h) => (
                <th key={h} className={`text-left text-[10px] uppercase tracking-widest ${t.muted} font-black px-2 py-2`}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.items.slice(0, 30).map((it, i) => {
              const high = it.infl > 20;
              return (
                <tr key={i} className={`border-b ${t.border}`}>
                  <td className="px-2 py-2 font-bold max-w-[200px] truncate">{it.name}</td>
                  <td className="px-2 py-2 num font-semibold">{inr(it.offline)}</td>
                  <td className="px-2 py-2 num font-semibold">{inr(it.online)}</td>
                  <td className={`px-2 py-2 num font-black ${high ? t.warn : t.pos}`}>
                    {it.infl >= 0 ? "+" : ""}{it.infl.toFixed(0)}%
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

// ============ CODB STEP ============
function CodbStep({ codb, setCodb, advanced, setAdvanced, breakdown, setBreakdown, t }) {
  const youGet = 100 - codb;
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <div className="flex items-baseline gap-3 mb-1">
        <div className={`text-xs font-black ${t.muted}`}>YOUR COSTS</div>
        <h2 className="text-lg md:text-xl font-black ultra-tight flex items-center gap-1">
          Cost of doing business <InfoTip t={t} text="Everything the platform takes per order: commission, ads, discounts, payment gateway, and GST on those fees. These are set by your contract, so enter your actual rates below." />
        </h2>
      </div>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>
        Enter the rates from your platform contract. Your total CODB is calculated for you - it's a fixed figure, not something to slide around.
      </p>

      {/* Individual cost inputs - always visible, since CODB is derived from these */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-5">
        {[
          { key: "commission", label: "Commission" }, { key: "ads", label: "Ads" },
          { key: "discount", label: "Discount" }, { key: "pg", label: "PG" }, { key: "gst", label: "GST on fees" },
        ].map((c) => (
          <div key={c.key} className={`border ${t.border} rounded-lg p-3 ${t.softBg}`}>
            <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1.5`}>{c.label}</div>
            <div className="flex items-baseline gap-0.5">
              <input type="number" step="0.01" value={breakdown[c.key]} min="0" max="50"
                onChange={(e) => setBreakdown({ ...breakdown, [c.key]: Math.max(0, Math.min(50, +e.target.value || 0)) })}
                className="w-full bg-transparent text-xl font-black ultra-tight num outline-none border-0 p-0" />
              <span className="text-sm font-black">%</span>
            </div>
          </div>
        ))}
      </div>

      {/* Computed CODB - read only */}
      <div className={`p-4 rounded-xl ${t.softBg} border ${t.border} flex flex-wrap items-center justify-between gap-3`}>
        <div>
          <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1`}>Your total cost of doing business</div>
          <div className="text-3xl md:text-4xl font-black num ultra-tight">{codb.toFixed(1)}<span className="text-xl">%</span></div>
        </div>
        <div className="text-right">
          <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-1`}>You keep, before food cost</div>
          <div className={`text-3xl md:text-4xl font-black num ultra-tight ${t.pos}`}>{youGet.toFixed(1)}<span className="text-xl">%</span></div>
        </div>
      </div>
    </section>
  );
}

// ============ TARGET MARGIN ============
function TargetMarginStep({ targetMargin, setTargetMargin, cuisineBenchmark, blendedMargin, t, dark }) {
  const SLIDER_MIN = 5, SLIDER_MAX = 45;
  const pos = (v) => ((Math.max(SLIDER_MIN, Math.min(SLIDER_MAX, v)) - SLIDER_MIN) / (SLIDER_MAX - SLIDER_MIN)) * 100;
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <div className="flex items-baseline gap-3 mb-1">
        <div className={`text-xs font-black ${t.muted}`}>YOUR GOAL</div>
        <h2 className="text-lg md:text-xl font-black ultra-tight">Healthy margin target</h2>
      </div>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>
        Items above this margin are your profit drivers - the ones worth promoting more.
        {cuisineBenchmark && <> Restaurants in your <span className="capitalize">{cuisineBenchmark.cuisine}</span> segment average about <strong className={t.text}>{cuisineBenchmark.avg.toFixed(0)}%</strong>.</>}
      </p>
      <div className="flex flex-wrap items-center gap-4 mb-2">
        <div className="flex items-baseline gap-1">
          <input type="number" value={targetMargin} min="5" max="50"
            onChange={(e) => setTargetMargin(Math.max(5, Math.min(50, +e.target.value || 0)))}
            className="w-24 bg-transparent text-5xl md:text-6xl font-black ultra-tight num outline-none border-0 p-0" />
          <span className="text-3xl md:text-4xl font-black">%</span>
        </div>
        <div className="flex-1 min-w-[200px]">
          {/* Slider with cuisine benchmark + current margin markers */}
          <div className="relative pt-5 pb-1">
            {cuisineBenchmark && (
              <div className="absolute top-0 -translate-x-1/2 flex flex-col items-center" style={{ left: pos(cuisineBenchmark.avg) + "%" }}>
                <span className={`text-[9px] font-black uppercase tracking-wider ${t.muted} whitespace-nowrap`}>Cuisine avg {cuisineBenchmark.avg.toFixed(0)}%</span>
                <span className={`w-px h-3 ${dark ? "bg-zinc-500" : "bg-zinc-400"}`} />
              </div>
            )}
            <input type="range" min={SLIDER_MIN} max={SLIDER_MAX} step={1} value={targetMargin}
              onChange={(e) => setTargetMargin(+e.target.value)} className="w-full" />
          </div>
          {blendedMargin > 0 && (
            <div className={`text-[11px] font-bold ${t.muted} mt-1`}>
              Your current blended margin: <span className={`num ${blendedMargin >= targetMargin ? t.pos : t.warn}`}>{blendedMargin.toFixed(1)}%</span>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

// ============ TOP PROFIT OPPORTUNITY CATEGORIES (expandable) ============
function OpportunityCategories({ categories, t, dark, targetMargin }) {
  const [expanded, setExpanded] = useState(categories[0]?.cat ?? null);
  const scoreLabel = (s) => (s >= 70 ? "Very High" : s >= 50 ? "High" : "Medium");
  const scoreTone = (s) => (s >= 70 ? t.pos : s >= 50 ? t.warn : t.muted);

  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <h2 className="text-lg md:text-xl font-black ultra-tight mb-1">Top profit opportunity categories</h2>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>
        Categories with healthy margins and room to grow. Promoting these - combos, ads, banners, better visibility - lifts your blended profitability <strong className={t.text}>without raising any prices.</strong> Tap a category to see its items. Breads and staples are excluded since they can't be promoted on their own.
      </p>

      {categories.length === 0 ? (
        <div className={`p-6 rounded-xl ${t.softBg} text-center text-sm font-bold ${t.muted}`}>
          No clear opportunity categories at a {targetMargin}% target yet. Try lowering your target margin, or upload your full menu.
        </div>
      ) : (
        <div className="space-y-2">
          {categories.map((c) => {
            const open = expanded === c.cat;
            return (
              <div key={c.cat} className={`border ${t.border} rounded-xl overflow-hidden`}>
                {/* Category header row */}
                <button onClick={() => setExpanded(open ? null : c.cat)}
                  className={`w-full flex items-center gap-3 p-4 text-left hover:${t.softBg} transition`}>
                  <div className="flex-1 min-w-0">
                    <div className="font-black ultra-tight text-base">{c.cat}</div>
                    <div className={`text-[11px] font-bold ${t.muted} mt-0.5`}>
                      {c.oppItems.length} item{c.oppItems.length > 1 ? "s" : ""} to push · avg {c.avgMargin.toFixed(0)}% margin · {inr(c.totalProfit)}/mo profit
                    </div>
                  </div>
                  <div className="text-right shrink-0">
                    <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black`}>Opportunity</div>
                    <div className={`text-sm font-black ${scoreTone(c.score)}`}>{scoreLabel(c.score)}</div>
                  </div>
                  <div className={`text-lg font-black ${t.muted} transition-transform ${open ? "rotate-180" : ""}`}>⌄</div>
                </button>

                {/* Expanded item table */}
                {open && (
                  <div className={`border-t ${t.border} ${t.softBg} p-3 md:p-4`}>
                    {/* Desktop */}
                    <div className="hidden md:block overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className={`border-b ${t.border}`}>
                            {["Item", "Margin", "Monthly Qty", "Monthly Profit", "Push?"].map((h) => (
                              <th key={h} className={`text-left text-[10px] uppercase tracking-widest ${t.muted} font-black px-2 py-2`}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {c.items.map((r) => {
                            const isOpp = c.oppItems.find((o) => o.idx === r.idx);
                            return (
                              <tr key={r.idx} className={`border-b ${t.border} last:border-0`}>
                                <td className="px-2 py-2.5 font-bold max-w-[220px] truncate">{r.name}</td>
                                <td className="px-2 py-2.5">
                                  <span className={`px-2 py-0.5 rounded-md text-xs font-black num border ${r.marginPct >= targetMargin ? `${t.posBg} ${t.pos}` : `${t.softBg} ${t.muted}`}`}>
                                    {r.marginPct.toFixed(0)}%
                                  </span>
                                </td>
                                <td className="px-2 py-2.5 num font-bold">{num(r.monthlyOrders)}</td>
                                <td className="px-2 py-2.5 num font-bold">{inr(r.monthlyProfit)}</td>
                                <td className="px-2 py-2.5">
                                  {isOpp
                                    ? <span className={`text-[11px] font-black ${t.pos}`}>✓ Promote</span>
                                    : <span className={`text-[11px] font-bold ${t.muted}`}>-</span>}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    {/* Mobile */}
                    <div className="md:hidden space-y-2">
                      {c.items.map((r) => {
                        const isOpp = c.oppItems.find((o) => o.idx === r.idx);
                        return (
                          <div key={r.idx} className={`flex items-center justify-between gap-2 p-2.5 rounded-lg ${t.card} border ${t.border}`}>
                            <div className="min-w-0">
                              <div className="font-bold truncate text-sm">{r.name}</div>
                              <div className={`text-[10px] font-bold ${t.muted}`}>{num(r.monthlyOrders)}/mo · {inr(r.monthlyProfit)}/mo</div>
                            </div>
                            <div className="text-right shrink-0">
                              <div className={`text-sm font-black num ${r.marginPct >= targetMargin ? t.pos : t.muted}`}>{r.marginPct.toFixed(0)}%</div>
                              {isOpp && <div className={`text-[10px] font-black ${t.pos}`}>✓ Promote</div>}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <div className={`text-[11px] font-bold ${t.muted} mt-3`}>
                      Promoting the ✓ items in {c.cat} can grow profit while keeping prices unchanged. Total category profit potential: {inr(c.totalProfit)}/mo at current volume.
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
}

// ============ PROMOTION IMPACT SIMULATOR ============
function PromotionSimulator({ opportunities, rows, rollup, codb, t, dark }) {
  // All promotable items are selectable (breads/staples excluded), grouped by category.
  const promotableRows = useMemo(() => rows.filter((r) => r.promotable), [rows]);
  const grouped = useMemo(() => {
    const g = {};
    promotableRows.forEach((r) => { (g[r.category] = g[r.category] || []).push(r); });
    Object.values(g).forEach((arr) => arr.sort((a, b) => b.marginPct - a.marginPct));
    return g;
  }, [promotableRows]);

  const [selIdx, setSelIdx] = useState(opportunities[0]?.idx ?? promotableRows[0]?.idx ?? null);
  const [addQty, setAddQty] = useState(50);

  useEffect(() => {
    if (promotableRows.length && !promotableRows.find((o) => o.idx === selIdx)) {
      setSelIdx(opportunities[0]?.idx ?? promotableRows[0].idx);
    }
  }, [promotableRows, opportunities]);

  const item = rows.find((r) => r.idx === selIdx);
  const impact = useMemo(() => {
    if (!item) return null;
    const addSales = item.price * addQty;
    const addPayout = item.grossPayout * addQty;
    const addProfit = item.net * addQty;
    const newRev = rollup.rev + addSales;
    const newProfit = rollup.profit + addProfit;
    const newMargin = newRev ? (newProfit / newRev) * 100 : 0;
    return {
      addSales, addPayout, addProfit,
      newMargin, oldMargin: rollup.blendedMargin,
      marginLift: newMargin - rollup.blendedMargin,
      yearlyProfit: addProfit * 12,
    };
  }, [item, addQty, rollup]);

  return (
    <section className={`border ${t.border} rounded-2xl p-5 md:p-6 ${dark ? "bg-zinc-900" : "bg-zinc-950 text-white"}`}>
      <div className={`text-xs font-black ${dark ? "text-lime-300" : "text-lime-400"} uppercase tracking-widest mb-2`}>Promotion impact simulator</div>
      <h2 className="text-xl md:text-2xl font-black ultra-tight mb-1">What if you sell more of a high-margin item?</h2>
      <p className={`text-sm ${dark ? "text-zinc-400" : "text-zinc-300"} font-medium mb-5`}>
        Pick any item from your menu and add extra monthly orders. See how your whole-restaurant profitability moves - no price changes involved. (Breads and staples are excluded - they can't be promoted on their own.)
      </p>

      <div className="grid sm:grid-cols-2 gap-4 mb-6">
        <div>
          <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-2`}>Choose any item to promote</div>
          <select value={selIdx ?? ""} onChange={(e) => setSelIdx(+e.target.value)}
            className={`w-full rounded-lg px-3 py-2.5 text-sm font-bold ${dark ? "bg-zinc-800 text-zinc-50 border-zinc-700" : "bg-white text-zinc-950 border-zinc-300"} border`}>
            {Object.keys(grouped).sort().map((cat) => (
              <optgroup key={cat} label={cat}>
                {grouped[cat].map((r) => (
                  <option key={r.idx} value={r.idx}>{r.name} - {r.marginPct.toFixed(0)}% margin · {num(r.monthlyOrders)}/mo</option>
                ))}
              </optgroup>
            ))}
          </select>
          {item && (
            <div className={`text-[11px] font-bold mt-2 ${dark ? "text-zinc-400" : "text-zinc-400"}`}>
              Currently {num(item.monthlyOrders)} orders/month at {inr(item.price)} · {item.marginPct.toFixed(0)}% margin
            </div>
          )}
        </div>
        <div>
          <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-2`}>Extra orders per month</div>
          <div className="flex items-baseline gap-1">
            <span className="text-xl font-black">+</span>
            <input type="number" value={addQty} min="0" max="2000"
              onChange={(e) => setAddQty(Math.max(0, +e.target.value || 0))}
              className="w-24 bg-transparent text-4xl font-black ultra-tight num outline-none border-0 p-0" />
          </div>
          <input type="range" min={0} max={300} step={5} value={Math.min(addQty, 300)}
            onChange={(e) => setAddQty(+e.target.value)} className="w-full mt-2" />
        </div>
      </div>

      {impact && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 md:gap-4 mb-5">
            <div>
              <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-1`}>Extra sales</div>
              <div className="text-2xl md:text-3xl font-black num ultra-tight">{inrCompact(impact.addSales)}</div>
            </div>
            <div>
              <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-1`}>Extra payout</div>
              <div className="text-2xl md:text-3xl font-black num ultra-tight">{inrCompact(impact.addPayout)}</div>
            </div>
            <div>
              <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-1`}>Extra profit</div>
              <div className="text-2xl md:text-3xl font-black num ultra-tight text-lime-400">{inrCompact(impact.addProfit)}</div>
            </div>
            <div>
              <div className={`text-[10px] uppercase tracking-widest ${dark ? "text-zinc-400" : "text-zinc-400"} font-black mb-1`}>Blended margin</div>
              <div className="text-2xl md:text-3xl font-black num ultra-tight">
                {impact.oldMargin.toFixed(1)}<span className="text-lime-400"> → {impact.newMargin.toFixed(1)}%</span>
              </div>
            </div>
          </div>
          <div className={`border-t ${dark ? "border-zinc-800" : "border-zinc-800"} pt-4 text-sm font-medium ${dark ? "text-zinc-300" : "text-zinc-300"}`}>
            {item && (
              <>If <strong className="text-white">{item.name}</strong> sells {addQty} more orders/month, your blended margin improves
              {" "}<strong className="text-lime-400">{impact.marginLift >= 0 ? "+" : ""}{impact.marginLift.toFixed(1)} points</strong>,
              adding <strong className="text-lime-400">{inrCompact(impact.addProfit)}/month</strong> ({inrCompact(impact.yearlyProfit)}/year) in profit - all without changing a single price.</>
            )}
          </div>
        </>
      )}
    </section>
  );
}

// ============ MARGIN MIX ============
function MarginMix({ categoryMix, t, dark, targetMargin }) {
  const maxProfitShare = Math.max(...categoryMix.map((c) => c.profitShare), 1);
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <h2 className="text-lg md:text-xl font-black ultra-tight mb-1">Where your profit comes from</h2>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>
        Categories that drive profit vs the ones that mostly drive volume. Push promotion toward the profit drivers.
      </p>
      <div className="space-y-3">
        {categoryMix.map((c) => {
          const efficient = c.profitShare > c.revShare + 2;
          const diluting = c.revShare > c.profitShare + 2;
          const tag = efficient ? "Profit driver" : diluting ? "Volume driver" : "Balanced";
          const tagTone = efficient ? `${t.pos} ${t.posBg}` : diluting ? `${t.warn} ${t.warnBg}` : `${t.muted} ${t.softBg}`;
          return (
            <div key={c.cat} className={`border ${t.border} rounded-xl p-4`}>
              <div className="flex items-center justify-between gap-2 mb-2">
                <div className="font-black ultra-tight">{c.cat}</div>
                <span className={`text-[10px] font-black uppercase tracking-wider px-2 py-0.5 rounded-md border ${tagTone}`}>{tag}</span>
              </div>
              <div className="flex items-center gap-4 mb-2">
                <div className="flex-1">
                  <div className={`flex justify-between text-[10px] font-bold ${t.muted} mb-1`}>
                    <span>Share of sales</span><span className="num">{c.revShare.toFixed(0)}%</span>
                  </div>
                  <div className={`h-1.5 rounded-full ${dark ? "bg-zinc-800" : "bg-zinc-200"}`}>
                    <div className={`h-full rounded-full ${dark ? "bg-zinc-500" : "bg-zinc-400"}`} style={{ width: c.revShare + "%" }} />
                  </div>
                </div>
                <div className="flex-1">
                  <div className={`flex justify-between text-[10px] font-bold ${t.muted} mb-1`}>
                    <span>Share of profit</span><span className="num">{c.profitShare.toFixed(0)}%</span>
                  </div>
                  <div className={`h-1.5 rounded-full ${dark ? "bg-zinc-800" : "bg-zinc-200"}`}>
                    <div className={`h-full rounded-full ${dark ? "bg-lime-400" : "bg-emerald-600"}`} style={{ width: c.profitShare + "%" }} />
                  </div>
                </div>
              </div>
              <div className={`text-[11px] font-bold ${t.muted}`}>
                {c.margin.toFixed(0)}% margin · {num(c.orders)} orders/mo
                {c.promotable === 0 && <span> · staple category, not for promotion</span>}
              </div>
            </div>
          );
        })}
      </div>
    </section>
  );
}

// ============ INSIGHTS ============
function InsightsView({ insights, t }) {
  return (
    <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
      <h2 className="text-lg md:text-xl font-black ultra-tight mb-1">Smart suggestions</h2>
      <p className={`text-sm ${t.muted} font-medium mb-5`}>Ways to grow profit through a better menu mix - never by raising prices.</p>
      <div className="space-y-3">
        {insights.map((s, i) => {
          const bg = s.tone === "opportunity" ? t.posBg : s.tone === "warning" ? t.warnBg : t.infoBg;
          const dot = s.tone === "opportunity" ? "bg-emerald-500" : s.tone === "warning" ? "bg-amber-500" : "bg-blue-500";
          return (
            <div key={i} className={`border rounded-xl p-4 ${bg} flex gap-3`}>
              <div className={`w-2 h-2 rounded-full ${dot} mt-2 shrink-0`} />
              <div>
                <div className="font-black text-sm md:text-base ultra-tight mb-1">{s.title}</div>
                <div className={`text-sm font-medium ${t.muted}`}>{s.body}</div>
              </div>
            </div>
          );
        })}
      </div>
    </section>
  );
}

// ============ INFO TOOLTIP ============
function InfoTip({ text, t }) {
  const [open, setOpen] = useState(false);
  return (
    <span className="relative inline-flex">
      <button type="button"
        onClick={(e) => { e.stopPropagation(); setOpen(!open); }}
        onMouseEnter={() => setOpen(true)} onMouseLeave={() => setOpen(false)}
        className={`w-3.5 h-3.5 rounded-full border ${t.border} ${t.muted} text-[9px] font-black inline-flex items-center justify-center hover:${t.text} shrink-0`}
        aria-label="info">i</button>
      {open && (
        <span className={`absolute z-50 bottom-full left-1/2 -translate-x-1/2 mb-1.5 w-48 p-2 rounded-lg border ${t.border} ${t.card} text-[11px] font-medium ${t.text} shadow-xl normal-case tracking-normal`} style={{ whiteSpace: "normal" }}>
          {text}
        </span>
      )}
    </span>
  );
}

// ============ UPLOAD MODAL (dual: offline + online, real OCR) ============
function UploadModal({ onClose, onImport, t, dark }) {
  const [offline, setOffline] = useState(null);   // parsed items
  const [online, setOnline] = useState(null);
  const [busy, setBusy] = useState(null);          // 'offline' | 'online' | null
  const [err, setErr] = useState(null);
  const offRef = useRef();
  const onRef = useRef();

  const handle = async (which, file) => {
    if (!file) return;
    setErr(null);
    setBusy(which);
    try {
      const items = await parseMenuFile(file);
      if (!items || items.length === 0) {
        setErr(`We couldn't read any items from that ${which} file. Try a clearer photo, or a PDF/Excel/CSV.`);
      } else {
        which === "offline" ? setOffline(items) : setOnline(items);
      }
    } catch (e) {
      console.error(e);
      setErr("Couldn't read the file: " + e.message);
    }
    setBusy(null);
  };

  const canImport = (offline && offline.length) || (online && online.length);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3" onClick={onClose}>
      <div className={`absolute inset-0 ${dark ? "bg-black/70" : "bg-zinc-900/30"} backdrop-blur-sm`} />
      <div className={`relative ${t.card} border ${t.border} rounded-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl`} onClick={(e) => e.stopPropagation()}>
        <div className={`sticky top-0 ${t.card} border-b ${t.border} px-5 py-3.5 flex items-center justify-between`}>
          <div className="text-base font-black ultra-tight">Upload your menu</div>
          <button onClick={onClose} className={`text-sm font-bold ${t.muted} hover:${t.text}`}>✕</button>
        </div>

        <div className="p-5">
          <p className={`text-sm ${t.muted} font-medium mb-4`}>
            Upload a photo, PDF, Word, Excel or CSV - we read it automatically. Add both menus to compare online vs dine-in pricing. Just one is fine too.
          </p>

          <div className="grid sm:grid-cols-2 gap-3 mb-4">
            <UploadBox
              label="Dine-in menu" sublabel="Your offline / in-restaurant prices"
              items={offline} busy={busy === "offline"}
              onPick={() => offRef.current?.click()} onClear={() => setOffline(null)} t={t} dark={dark}
            />
            <UploadBox
              label="Online menu" sublabel="Your delivery app prices"
              items={online} busy={busy === "online"}
              onPick={() => onRef.current?.click()} onClear={() => setOnline(null)} t={t} dark={dark}
            />
          </div>

          <input ref={offRef} type="file" className="hidden"
            accept=".csv,.xlsx,.xls,.docx,.pdf,.png,.jpg,.jpeg,.webp,.bmp,.gif"
            onChange={(e) => { handle("offline", e.target.files?.[0]); e.target.value = ""; }} />
          <input ref={onRef} type="file" className="hidden"
            accept=".csv,.xlsx,.xls,.docx,.pdf,.png,.jpg,.jpeg,.webp,.bmp,.gif"
            onChange={(e) => { handle("online", e.target.files?.[0]); e.target.value = ""; }} />

          {busy && (
            <div className={`p-3 rounded-lg ${t.softBg} text-sm font-bold ${t.muted} mb-3`}>
              Reading your {busy} menu… photos take a few seconds while we recognise the text.
            </div>
          )}
          {err && (
            <div className={`p-3 rounded-lg border ${t.border} ${dark ? "bg-rose-950/40" : "bg-rose-50"} text-sm font-bold ${t.neg} mb-3`}>
              {err}
            </div>
          )}

          <button
            disabled={!canImport}
            onClick={() => onImport({ online, offline })}
            className={`w-full text-sm font-black px-4 py-3 rounded-lg ${canImport ? t.accent : `${t.softBg} ${t.muted} cursor-not-allowed`}`}>
            {canImport ? "Show my dashboard" : "Upload at least one menu"}
          </button>
          <div className={`text-[11px] ${t.muted} font-medium mt-3 text-center`}>
            We match your menu to your restaurant automatically and pull in your sales picture.
          </div>
        </div>
      </div>
    </div>
  );
}

function UploadBox({ label, sublabel, items, busy, onPick, onClear, t, dark }) {
  const done = items && items.length > 0;
  return (
    <div className={`border-2 border-dashed rounded-xl p-4 ${done ? (dark ? "border-lime-500/50 bg-lime-950/10" : "border-emerald-400 bg-emerald-50/40") : t.border}`}>
      <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mb-0.5`}>{label}</div>
      <div className={`text-[11px] ${t.muted} font-semibold mb-3`}>{sublabel}</div>
      {done ? (
        <div>
          <div className={`text-2xl font-black ultra-tight ${t.pos}`}>{items.length} items</div>
          <div className={`text-[11px] ${t.muted} font-bold mt-0.5 truncate`}>
            {items.slice(0, 3).map((i) => i.name).join(", ")}{items.length > 3 ? "…" : ""}
          </div>
          <div className="flex gap-2 mt-2">
            <button onClick={onPick} className={`text-[11px] font-bold ${t.muted} hover:${t.text}`}>Replace</button>
            <button onClick={onClear} className={`text-[11px] font-bold ${t.muted} hover:${t.neg}`}>Remove</button>
          </div>
        </div>
      ) : (
        <button onClick={onPick} disabled={busy}
          className={`w-full py-4 rounded-lg ${t.softBg} text-sm font-black hover:opacity-80 ${busy ? "opacity-50" : ""}`}>
          {busy ? "Reading…" : "Choose file or photo"}
        </button>
      )}
    </div>
  );
}

// ============ KAM ADMIN PANEL ============
function AdminPanel({ t, dark, onExit, onThemeToggle }) {
  const [salesFile, setSalesFile] = useState(null);
  const [menuMixFile, setMenuMixFile] = useState(null);
  const [salesStats, setSalesStats] = useState(null);
  const [mixStats, setMixStats] = useState(null);
  const [busy, setBusy] = useState(null);
  const salesRef = useRef();
  const mixRef = useRef();

  const handleSales = async (file) => {
    if (!file) return;
    setBusy("sales");
    try {
      const rows = await readTabular(file);
      // Validate expected Sales columns
      const cols = rows[0] ? Object.keys(rows[0]).map((c) => c.toLowerCase()) : [];
      const hasReq = ["restaurant_id", "restaurant_name", "gmv", "net_sales", "area"].filter((c) => cols.some((x) => x.includes(c)));
      const restaurants = new Set(rows.map((r) => r.restaurant_name || r.Restaurant_Name).filter(Boolean));
      setSalesStats({
        rows: rows.length,
        restaurants: restaurants.size,
        matched: hasReq.length,
        expected: 5,
        sample: [...restaurants].slice(0, 5),
      });
      setSalesFile(file.name);
    } catch (e) {
      setSalesStats({ error: e.message });
    }
    setBusy(null);
  };

  const handleMix = async (file) => {
    if (!file) return;
    setBusy("mix");
    try {
      const rows = await readTabular(file);
      const cols = rows[0] ? Object.keys(rows[0]).map((c) => c.toLowerCase()) : [];
      const hasReq = ["restaurant_id", "item_name", "category", "orders", "unit_price"].filter((c) => cols.some((x) => x.includes(c)));
      const items = new Set(rows.map((r) => r.item_name || r.Item_Name).filter(Boolean));
      setMixStats({
        rows: rows.length,
        items: items.size,
        matched: hasReq.length,
        expected: 5,
      });
      setMenuMixFile(file.name);
    } catch (e) {
      setMixStats({ error: e.message });
    }
    setBusy(null);
  };

  return (
    <div className={`min-h-screen ${t.bg} ${t.text} antialiased`}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        * { font-family: 'Inter', system-ui, sans-serif; }
        .num { font-variant-numeric: tabular-nums; }
        .ultra-tight { letter-spacing: -0.05em; }
      `}</style>

      <MockBanner />

      <header className={`border-b ${t.border} sticky top-0 z-30 ${t.bg}`}>
        <div className="max-w-4xl mx-auto px-4 md:px-6 h-14 flex items-center gap-3">
          <div className={`px-2 py-0.5 rounded-md text-[10px] font-black uppercase tracking-widest ${t.accent}`}>KAM Admin</div>
          <div className="font-black text-base tight">Backend data uploads</div>
          <button onClick={onThemeToggle} className={`ml-auto text-xs font-bold px-2 py-1.5 rounded-md border ${t.border} hover:${t.softBg}`}>
            {dark ? "☀" : "☾"}
          </button>
          <button onClick={onExit} className={`text-xs font-bold px-3 py-1.5 rounded-md border ${t.border} hover:${t.softBg}`}>
            ← Partner view
          </button>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 md:px-6 py-8 space-y-6">
        <section>
          <div className={`text-[11px] uppercase tracking-widest ${t.muted} font-black mb-2`}>Internal · Account Managers only</div>
          <h1 className="text-2xl md:text-3xl font-black ultra-tight">Upload source data</h1>
          <p className={`text-sm ${t.muted} font-medium mt-2 max-w-2xl`}>
            This is where the KAM team uploads the raw backend files that power every partner dashboard: the sales data export and the Menu Mix export. Partners never see this screen or the raw data - they only see anonymized, aggregated benchmarks.
          </p>
        </section>

        <div className={`rounded-xl border ${t.border} ${dark ? "bg-amber-950/30" : "bg-amber-50"} p-4 text-sm font-bold ${t.warn}`}>
          In this demo build, uploaded files are parsed and validated in your browser only - nothing is stored or sent anywhere. In production this would write to the secure multi-tenant backend with role-based access.
        </div>

        {/* Sales upload */}
        <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
          <div className="flex items-baseline gap-3 mb-1">
            <div className={`text-xs font-black ${t.muted}`}>FILE 1</div>
            <h2 className="text-lg font-black ultra-tight">Sales data export</h2>
          </div>
          <p className={`text-sm ${t.muted} font-medium mb-4`}>
            Restaurant-level sales report (CSV or Excel). Expected columns include restaurant_id, restaurant_name, area, gmv, net_sales, orders and discounts.
          </p>
          <input ref={salesRef} type="file" accept=".csv,.xlsx,.xls" className="hidden"
            onChange={(e) => { handleSales(e.target.files?.[0]); e.target.value = ""; }} />
          <button onClick={() => salesRef.current?.click()} disabled={busy === "sales"}
            className={`text-sm font-bold px-4 py-2.5 rounded-lg ${t.accent} ${busy === "sales" ? "opacity-50" : ""}`}>
            {busy === "sales" ? "Reading…" : "Upload sales data"}
          </button>
          {salesFile && salesStats && !salesStats.error && (
            <div className={`mt-4 p-4 rounded-xl ${t.softBg} border ${t.border}`}>
              <div className={`text-sm font-black mb-2 ${t.pos}`}>✓ {salesFile} loaded</div>
              <div className="grid grid-cols-3 gap-3 text-center">
                <AdminStat t={t} label="Rows" value={num(salesStats.rows)} />
                <AdminStat t={t} label="Restaurants" value={num(salesStats.restaurants)} />
                <AdminStat t={t} label="Schema match" value={`${salesStats.matched}/${salesStats.expected}`} />
              </div>
            </div>
          )}
          {salesStats?.error && <div className={`mt-3 text-sm font-bold ${t.neg}`}>{salesStats.error}</div>}
        </section>

        {/* Menu Mix upload */}
        <section className={`border ${t.border} ${t.card} rounded-2xl p-5 md:p-6`}>
          <div className="flex items-baseline gap-3 mb-1">
            <div className={`text-xs font-black ${t.muted}`}>FILE 2</div>
            <h2 className="text-lg font-black ultra-tight">Menu Mix export</h2>
          </div>
          <p className={`text-sm ${t.muted} font-medium mb-4`}>
            Item-level menu performance (CSV or Excel). Expected columns include restaurant_id, item_name, category, orders, quantity and unit_price.
          </p>
          <input ref={mixRef} type="file" accept=".csv,.xlsx,.xls" className="hidden"
            onChange={(e) => { handleMix(e.target.files?.[0]); e.target.value = ""; }} />
          <button onClick={() => mixRef.current?.click()} disabled={busy === "mix"}
            className={`text-sm font-bold px-4 py-2.5 rounded-lg ${t.accent} ${busy === "mix" ? "opacity-50" : ""}`}>
            {busy === "mix" ? "Reading…" : "Upload Menu Mix"}
          </button>
          {menuMixFile && mixStats && !mixStats.error && (
            <div className={`mt-4 p-4 rounded-xl ${t.softBg} border ${t.border}`}>
              <div className={`text-sm font-black mb-2 ${t.pos}`}>✓ {menuMixFile} loaded</div>
              <div className="grid grid-cols-3 gap-3 text-center">
                <AdminStat t={t} label="Rows" value={num(mixStats.rows)} />
                <AdminStat t={t} label="Unique items" value={num(mixStats.items)} />
                <AdminStat t={t} label="Schema match" value={`${mixStats.matched}/${mixStats.expected}`} />
              </div>
            </div>
          )}
          {mixStats?.error && <div className={`mt-3 text-sm font-bold ${t.neg}`}>{mixStats.error}</div>}
        </section>

        {salesFile && menuMixFile && (
          <div className={`rounded-xl border ${t.border} ${t.posBg} p-5`}>
            <div className="font-black text-lg ultra-tight mb-1">Both files loaded</div>
            <div className={`text-sm font-medium ${t.muted}`}>
              In production, the system would now join these on <strong className={t.text}>restaurant_id</strong>, anonymize competitor data, rebuild the benchmark engine, and refresh every partner dashboard. Partners only ever see aggregated, anonymized benchmarks - never another restaurant's raw numbers.
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function AdminStat({ t, label, value }) {
  return (
    <div>
      <div className={`text-xl font-black num ultra-tight`}>{value}</div>
      <div className={`text-[10px] uppercase tracking-widest ${t.muted} font-black mt-0.5`}>{label}</div>
    </div>
  );
}

async function readTabular(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  if (ext === "csv") {
    return new Promise((resolve, reject) => {
      Papa.parse(file, { header: true, skipEmptyLines: "greedy", complete: (r) => resolve(r.data), error: reject });
    });
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
}