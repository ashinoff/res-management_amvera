require('dotenv').config();
const { Sequelize, DataTypes, Op } = require('sequelize');
const cloudinary = require('cloudinary').v2;
const fetch = require('node-fetch');

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

const sequelize = new Sequelize(process.env.DATABASE_URL, {
  dialect: 'postgres',
  dialectOptions: {
    ssl: { require: true, rejectUnauthorized: false }
  },
  logging: false
});

const CheckHistory = sequelize.define('CheckHistory', {
  id: { type: DataTypes.INTEGER, primaryKey: true },
  attachments: { type: DataTypes.JSON, defaultValue: [] }
});

async function fixPdfs() {
  try {
    await sequelize.authenticate();
    console.log('✅ Connected to database\n');
    
    const records = await CheckHistory.findAll({
      where: {
        attachments: { [Op.ne]: [] }
      }
    });
    
    console.log(`Found ${records.length} records with attachments\n`);
    
    let fixedCount = 0;
    let errorCount = 0;
    
    for (const record of records) {
      if (!record.attachments) continue;
      
      let needsUpdate = false;
      const newAttachments = [];
      
      for (const file of record.attachments) {
        const isPdf = file.original_name && 
                     file.original_name.toLowerCase().endsWith('.pdf');
        
        if (!isPdf) {
          newAttachments.push(file);
          continue;
        }
        
        console.log(`\n📄 Checking: ${file.original_name}`);
        console.log(`   Public ID: ${file.public_id}`);
        console.log(`   URL: ${file.url}`);
        
        try {
          // Проверяем как RAW
          try {
            await cloudinary.api.resource(file.public_id, { 
              resource_type: 'raw' 
            });
            console.log(`   ✅ Already RAW - OK!`);
            newAttachments.push(file);
            continue;
          } catch (e) {
            // Не найден как raw
          }
          
          // Проверяем как IMAGE
          let imageResource;
          try {
            imageResource = await cloudinary.api.resource(file.public_id, { 
              resource_type: 'image' 
            });
            console.log(`   ⚠️ Found as IMAGE - needs fixing!`);
          } catch (e) {
            console.log(`   ❌ Not found in Cloudinary`);
            newAttachments.push(file);
            errorCount++;
            continue;
          }
          
          // Скачиваем файл
          console.log(`   📥 Downloading...`);
          const response = await fetch(file.url);
          const buffer = Buffer.from(await response.arrayBuffer());
          
          // Загружаем заново как RAW
          console.log(`   📤 Re-uploading as RAW...`);
          
          const result = await new Promise((resolve, reject) => {
            const uploadStream = cloudinary.uploader.upload_stream(
              {
                folder: 'res-management',
                resource_type: 'raw',
                public_id: file.public_id + '_fixed',
                access_mode: 'public',
                overwrite: false
              },
              (error, result) => {
                if (error) reject(error);
                else resolve(result);
              }
            );
            
            const bufferStream = require('stream').Readable.from(buffer);
            bufferStream.pipe(uploadStream);
          });
          
          // Удаляем старый файл (как image)
          console.log(`   🗑️ Deleting old version...`);
          await cloudinary.uploader.destroy(file.public_id, { 
            resource_type: 'image' 
          });
          
          // Обновляем запись
          newAttachments.push({
            ...file,
            url: result.secure_url,
            public_id: result.public_id,
            resource_type: 'raw'
          });
          
          needsUpdate = true;
          fixedCount++;
          console.log(`   ✅ FIXED! New URL: ${result.secure_url}`);
          
        } catch (error) {
          console.error(`   ❌ Error: ${error.message}`);
          newAttachments.push(file);
          errorCount++;
        }
      }
      
      if (needsUpdate) {
        await record.update({ attachments: newAttachments });
        console.log(`   💾 Database updated for record ${record.id}`);
      }
    }
    
    console.log(`\n${'='.repeat(50)}`);
    console.log(`✅ Migration complete!`);
    console.log(`   Fixed: ${fixedCount} PDFs`);
    console.log(`   Errors: ${errorCount}`);
    console.log(`${'='.repeat(50)}\n`);
    
    process.exit(0);
    
  } catch (error) {
    console.error('\n❌ Migration failed:', error);
    process.exit(1);
  }
}

fixPdfs();
