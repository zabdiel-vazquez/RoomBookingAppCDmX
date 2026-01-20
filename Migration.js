/**
 * Migration Script
 * ================
 * Run these functions ONCE to migrate from insecure to secure configuration
 *
 * INSTRUCTIONS:
 * 1. Open Apps Script Editor
 * 2. Add this file to your project
 * 3. Run `migrateAllSecuritySettings()` from the editor
 * 4. Verify Script Properties were created
 * 5. Update bot.js and App.js to use Config.js functions
 * 6. Delete this file after successful migration
 */

/**
 * Main migration function - run this once
 */
function migrateAllSecuritySettings() {
  console.log('🔒 Starting security migration...\n');

  try {
    // Step 1: Migrate Slack user mappings
    console.log('Step 1: Migrating Slack user mappings...');
    var mappingCount = migrateSlackUserMappings();
    console.log('✅ Migrated ' + mappingCount + ' Slack user mappings\n');

    // Step 2: Set basic config
    console.log('Step 2: Setting up Script Properties...');
    setupScriptProperties();
    console.log('✅ Script Properties configured\n');

    // Step 3: Verify migration
    console.log('Step 3: Verifying migration...');
    var verification = verifyMigration();

    if (verification.success) {
      console.log('✅ Migration completed successfully!\n');
      console.log('NEXT STEPS:');
      console.log('1. Review Script Properties in Project Settings');
      console.log('2. Update SLACK_BOT_TOKEN with your actual token');
      console.log('3. Test the application with secure config');
      console.log('4. Update App.js and bot.js to use Config.js functions');
      console.log('5. Delete this Migration.js file');
    } else {
      console.error('❌ Migration verification failed:');
      verification.errors.forEach(function(err) {
        console.error('  - ' + err);
      });
    }

  } catch (error) {
    console.error('❌ Migration failed:', error.message);
    console.error('Stack:', error.stack);
  }
}

/**
 * Migrate hardcoded Slack user mappings to Script Properties
 */
function migrateSlackUserMappings() {
  // COPY-PASTE the DEFAULT_SLACK_USER_OVERRIDES object from bot.js here
  // This is the LAST time these should appear in code
  var legacyMappings = {
    'savvy@apollo.io': 'U020NN8Q1GU',
    'juan.nieto@apollo.io': 'U09BNNPR32R',
    'erika.barrios@apollo.io': 'U08RNAQBK60',
    'alvaro.cabrera@apollo.io': 'U09G55MB3MX',
    'lupita.garcia@apollo.io': 'U092J6KJT6G',
    'isabela.ayala@apollo.io': 'U096B5HUWRC',
    'yaqueline.murillo@apollo.io': 'U09C23VTLMR',
    'gabriela.gonzalez@apollo.io': 'U08PBPE1DNZ',
    'carolina.servin@apollo.io': 'U097BUXD7QW',
    'paola.dominguez@apollo.io': 'U09FH0ETXE5',
    'daniel.camino@apollo.io': 'U094DT0CSJN',
    'mario.portillo@apollo.io': 'U09D1A80WLK',
    'guillermo.diaz@apollo.io': 'U09CAKK1Z2V',
    'francisco.ruiz@apollo.io': 'U09CAKU1SHF',
    'israel.aguirre@apollo.io': 'U09HFL3LBAQ',
    'mario.perez@apollo.io': 'U09CWNP2P3X',
    'santiago.hernandez@apollo.io': 'U09CH7H5N8H',
    'pedro.chavero@apollo.io': 'U09HEEYFNSU',
    'alan.gonzalez@apollo.io': 'U09HA4VHB49',
    'hector.mejia@apollo.io': 'U09EF1M6CU2',
    'emiliano.mejia@apollo.io': 'U09H65T370Q',
    'irvin.medina@apollo.io': 'U09GMUPDDCA',
    'jesus.guzman@apollo.io': 'U09G5HWQNUB',
    'omar.camarena@apollo.io': 'U09FZ1XNV75',
    'fernando.leal@apollo.io': 'U09G2L8Q04S',
    'christian.zamora@apollo.io': 'U097QQ6NVFK',
    'edin.martinez@apollo.io': 'U09FFEHG2NW',
    'alejandro.castro@apollo.io': 'U09FFEL854M',
    'rodrigo.lopez@apollo.io': 'U09GD4MRNK9',
    'francisco.gonzalez@apollo.io': 'U09G70ZHSGD',
    'miguel.avila@apollo.io': 'U09FTVD4J2U',
    'diego.alvarez@apollo.io': 'U09HAHQH4U8',
    'gabriel.morales@apollo.io': 'U09H9HF98DF',
    'juan.cardenas@apollo.io': 'U09HJ80QDHJ',
    'hugo.villegas@apollo.io': 'U09H9QF70JL',
    'jose.luna@apollo.io': 'U09GFGMJG5E',
    'julio.delgado@apollo.io': 'U096BQHTGLL',
    'ivan.garcia@apollo.io': 'U09DYQL4FHL',
    'erick.sanchez@apollo.io': 'U0987SG83AX',
    'carlos.jimenez@apollo.io': 'U09D30D2NGE',
    'hugo.flores@apollo.io': 'U09CKQK77RD',
    'jorge.ramirez@apollo.io': 'U09GM9FP86S',
    'arturo.moreno@apollo.io': 'U09HKP4FVPL',
    'jorge.velazquez@apollo.io': 'U09CA4S6TCN',
    'edgardo.martinez@apollo.io': 'U09HJGTAXD4',
    'rodrigo.miranda@apollo.io': 'U09HALXEU0P',
    'paul.mendez@apollo.io': 'U09G57P6P83',
    'juan.guerrero@apollo.io': 'U09GU4V56P9',
    'zabdiel.vazquez@apollo.io': 'U08L2CVG29W',
    'manuel.orozco@apollo.io': 'U09HKSP7Z69',
    'salvador.contreras@apollo.io': 'U09GA2RLSAJ',
    'jose.reyes@apollo.io': 'U09H6NKRWAJ',
    'luis.cervantes@apollo.io': 'U09CNUCPFB3'
  };

  var count = 0;
  var props = PropertiesService.getScriptProperties();

  for (var email in legacyMappings) {
    if (legacyMappings.hasOwnProperty(email)) {
      try {
        var slackId = legacyMappings[email];
        var key = 'SLACK_USER_MAP_' + email.toLowerCase();

        props.setProperty(key, slackId);
        count++;

        // Log progress every 10 users (without exposing emails)
        if (count % 10 === 0) {
          console.log('  ... migrated ' + count + ' users');
        }
      } catch (error) {
        console.error('Failed to migrate ' + email + ':', error.message);
      }
    }
  }

  return count;
}

/**
 * Setup initial Script Properties configuration
 */
function setupScriptProperties() {
  var props = PropertiesService.getScriptProperties();

  // Set default values (you'll need to update these manually)
  var defaultConfig = {
    'SLACK_BOT_TOKEN': 'REPLACE_WITH_YOUR_TOKEN',
    'SLACK_ADMIN_ID': 'U08L2CVG29W', // Zabdiel's ID from bot.js
    'SLACK_DEFAULT_CHANNEL': 'C09GZJSPDV4', // From bot.js
    'ADMIN_EMAILS': 'zabdiel.vazquez@apollo.io,it@apollo.io',
    'WEB_APP_URL': '' // Will be set after deployment
  };

  var existing = {};
  var allProps = props.getProperties();

  // Don't overwrite existing properties
  for (var key in defaultConfig) {
    if (defaultConfig.hasOwnProperty(key)) {
      if (!allProps[key]) {
        props.setProperty(key, defaultConfig[key]);
        console.log('  Set ' + key + ' = ' + (key.includes('TOKEN') ? '[REDACTED]' : defaultConfig[key]));
      } else {
        console.log('  Skipped ' + key + ' (already exists)');
        existing[key] = true;
      }
    }
  }

  if (Object.keys(existing).length > 0) {
    console.log('\n⚠️  Some properties already existed and were not overwritten.');
    console.log('   Review Script Properties to ensure values are correct.');
  }
}

/**
 * Verify migration was successful
 */
function verifyMigration() {
  var errors = [];
  var warnings = [];
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();

  // Check required Script Properties exist
  var requiredProps = ['SLACK_BOT_TOKEN', 'SLACK_ADMIN_ID', 'SLACK_DEFAULT_CHANNEL'];

  requiredProps.forEach(function(key) {
    if (!allProps[key]) {
      errors.push('Missing required property: ' + key);
    } else if (allProps[key] === 'REPLACE_WITH_YOUR_TOKEN') {
      warnings.push('Property ' + key + ' has placeholder value - update it manually');
    }
  });

  // Check Slack user mappings were migrated
  var mappingCount = 0;
  for (var key in allProps) {
    if (allProps.hasOwnProperty(key) && key.startsWith('SLACK_USER_MAP_')) {
      mappingCount++;
    }
  }

  if (mappingCount === 0) {
    errors.push('No Slack user mappings found - migration may have failed');
  } else {
    console.log('  Found ' + mappingCount + ' Slack user mappings');
  }

  // Check if Config.js functions are available
  try {
    if (typeof getSecureConfig === 'function') {
      var config = getSecureConfig();
      console.log('  Config.js functions are available ✓');
    } else {
      errors.push('Config.js functions not found - did you add Config.js to the project?');
    }
  } catch (error) {
    errors.push('Error testing Config.js: ' + error.message);
  }

  return {
    success: errors.length === 0,
    errors: errors,
    warnings: warnings,
    stats: {
      slackMappings: mappingCount,
      scriptProperties: Object.keys(allProps).length
    }
  };
}

/**
 * Rollback migration (use only if something went wrong)
 */
function rollbackMigration() {
  var confirmation = Browser.msgBox(
    'Rollback Migration',
    'This will DELETE all Script Properties created by migration. Continue?',
    Browser.Buttons.YES_NO
  );

  if (confirmation !== 'yes') {
    console.log('Rollback cancelled');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var deleted = 0;

  // Delete Slack user mappings
  for (var key in allProps) {
    if (allProps.hasOwnProperty(key) && key.startsWith('SLACK_USER_MAP_')) {
      props.deleteProperty(key);
      deleted++;
    }
  }

  // Delete config properties
  var configKeys = ['SLACK_BOT_TOKEN', 'SLACK_ADMIN_ID', 'SLACK_DEFAULT_CHANNEL', 'ADMIN_EMAILS', 'WEB_APP_URL'];
  configKeys.forEach(function(key) {
    if (allProps[key]) {
      props.deleteProperty(key);
      deleted++;
    }
  });

  console.log('Rollback complete. Deleted ' + deleted + ' properties.');
}

/**
 * Export current Script Properties for backup
 */
function exportScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();

  // Redact sensitive values
  var sanitized = {};
  for (var key in allProps) {
    if (allProps.hasOwnProperty(key)) {
      if (key.includes('TOKEN') || key.includes('PASSWORD')) {
        sanitized[key] = '[REDACTED]';
      } else {
        sanitized[key] = allProps[key];
      }
    }
  }

  console.log('Script Properties Backup:');
  console.log(JSON.stringify(sanitized, null, 2));

  return sanitized;
}

/**
 * Test Slack user lookup after migration
 */
function testSlackUserLookup() {
  var testEmails = [
    'zabdiel.vazquez@apollo.io',
    'juan.nieto@apollo.io',
    'nonexistent@apollo.io'
  ];

  console.log('Testing Slack user lookup:\n');

  testEmails.forEach(function(email) {
    try {
      var slackId = getSlackUserMapping(email);
      if (slackId) {
        console.log('✅ ' + email + ' → ' + slackId);
      } else {
        console.log('❌ ' + email + ' → not found');
      }
    } catch (error) {
      console.error('❌ ' + email + ' → error: ' + error.message);
    }
  });
}

/**
 * Cleanup old cache entries (run periodically)
 */
function cleanupOldCacheEntries() {
  var cache = CacheService.getScriptCache();

  // CacheService doesn't provide a way to list all keys
  // so we can only clear specific patterns we know about

  console.log('Cache cleanup completed.');
  console.log('Note: Cache entries auto-expire based on TTL.');
}

/**
 * Generate migration report
 */
function generateMigrationReport() {
  console.log('======================================');
  console.log('SECURITY MIGRATION REPORT');
  console.log('======================================\n');

  console.log('Date:', new Date().toISOString());
  console.log('User:', Session.getActiveUser().getEmail());
  console.log('\n');

  var verification = verifyMigration();

  console.log('STATUS:', verification.success ? '✅ SUCCESS' : '❌ FAILED');
  console.log('\n');

  console.log('STATISTICS:');
  console.log('  - Slack user mappings:', verification.stats.slackMappings);
  console.log('  - Total Script Properties:', verification.stats.scriptProperties);
  console.log('\n');

  if (verification.errors.length > 0) {
    console.log('ERRORS:');
    verification.errors.forEach(function(err) {
      console.log('  ❌ ' + err);
    });
    console.log('\n');
  }

  if (verification.warnings.length > 0) {
    console.log('WARNINGS:');
    verification.warnings.forEach(function(warn) {
      console.log('  ⚠️  ' + warn);
    });
    console.log('\n');
  }

  console.log('======================================');

  return verification;
}
