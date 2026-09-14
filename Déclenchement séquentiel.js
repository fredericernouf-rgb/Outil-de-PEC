/****************************************************
 * SÉQUENCE (Script 4) - VERSION ASYNCHRONE
 * Compatible AppSheet (timeout 120s)
 *
 * FONCTIONNEMENT :
 * 1. AppSheet appelle lancerRapportComplet()
 * 2. Réponse immédiate "Traitement lancé"
 * 3. Trigger démarre Script 1 après 10 secondes
 * 4. Puis Script 3 + 2 après 3 minutes
 *
 * AMÉLIORATIONS :
 * ✅ Réponse instantanée à AppSheet
 * ✅ Exécution asynchrone via triggers
 * ✅ Logs structurés complets
 ****************************************************/

/****************************************************
 * ═══════════════════════════════════════════════
 * FONCTION APPELÉE PAR APPSHEET (Réponse rapide)
 * ═══════════════════════════════════════════════
 ****************************************************/
function lancerRapportComplet() {
  var executionId = Utilities.getUuid();
  
  logStructured_('INFO', 'ORCHESTRATION', '🚀 Demande reçue depuis AppSheet', { 
    executionId: executionId 
  });
  
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 ORCHESTRATION - Demande AppSheet reçue');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🆔 Execution ID : ' + executionId);

  // ✅ Sauvegarde l'executionId pour traçabilité
  PropertiesService.getScriptProperties().setProperty('lastExecutionId', executionId);

  // ✅ Nettoyage triggers existants
  cleanupExistingTriggers_();

  // ✅ Programme le démarrage de la séquence dans 10 secondes
  try {
    ScriptApp.newTrigger('demarrerSequenceAsynchrone_')
      .timeBased()
      .after(10 * 1000) // 10 secondes
      .create();

    logStructured_('INFO', 'ORCHESTRATION', '✅ Séquence programmée', {
      executionId: executionId,
      delaySeconds: 10
    });

    Logger.log('✅ Séquence programmée (démarrage dans 10 secondes)');
    Logger.log('═══════════════════════════════════════════════════');

    // ✅ RÉPONSE IMMÉDIATE À APPSHEET
    return {
      success: true,
      message: 'Traitement lancé avec succès',
      executionId: executionId,
      estimatedDuration: '5-10 minutes',
      status: 'En cours'
    };

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur programmation', {
      executionId: executionId,
      error: e.toString()
    });

    Logger.log('❌ ERREUR : ' + e.toString());

    return {
      success: false,
      message: 'Erreur lors du lancement',
      error: e.toString()
    };
  }
}

/****************************************************
 * ═══════════════════════════════════════════════
 * PHASE 1 : DELTA + SCRIPT 1 (Trigger automatique)
 * ═══════════════════════════════════════════════
 ****************************************************/
function demarrerSequenceAsynchrone_() {
  var startTime = Date.now();
  var executionId = PropertiesService.getScriptProperties().getProperty('lastExecutionId') || Utilities.getUuid();

  // ✅ Supprime ce trigger
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'demarrerSequenceAsynchrone_') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  logStructured_('INFO', 'ORCHESTRATION', '🚀 Démarrage Phase 1 asynchrone', { 
    executionId: executionId 
  });
  
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 ORCHESTRATION - PHASE 1 (Asynchrone)');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🆔 Execution ID : ' + executionId);
  Logger.log('');

  // ═══════════════════════════════════════════════
  // ÉTAPE 0 : Calcul DELTA
  // ═══════════════════════════════════════════════
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 ÉTAPE 0 : CALCUL DELTA');
  Logger.log('═══════════════════════════════════════════════════');

  var todoKeys;
  try {
    todoKeys = computeAndSaveDelta_();
    
    logStructured_('INFO', 'ORCHESTRATION', '✅ DELTA calculé', {
      executionId: executionId,
      count: todoKeys.length,
      clients: todoKeys.join(', ')
    });

    if (todoKeys.length === 0) {
      Logger.log('✅ Aucun changement détecté, rien à faire.');
      Logger.log('═══════════════════════════════════════════════════');
      
      logStructured_('INFO', 'ORCHESTRATION', '✅ Terminé (aucun changement)', {
        executionId: executionId,
        duration: ((Date.now() - startTime) / 1000).toFixed(1) + 's'
      });
      
      return;
    }

    Logger.log('🎯 ' + todoKeys.length + ' client(s) à traiter : ' + todoKeys.join(', '));
    Logger.log('');

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur calcul DELTA', {
      executionId: executionId,
      error: e.toString(),
      stack: e.stack || ''
    });
    Logger.log('❌ ERREUR CRITIQUE : Impossible de calculer le DELTA');
    Logger.log(e.toString());
    return; // ✅ Arrêt propre au lieu de throw
  }

  // ═══════════════════════════════════════════════
  // ÉTAPE 1 : Création fichiers (Script 1)
  // ═══════════════════════════════════════════════
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📁 ÉTAPE 1 : CRÉATION FICHIERS (Script 1)');
  Logger.log('═══════════════════════════════════════════════════');

  try {
    createRapportClientFiles();
    
    logStructured_('INFO', 'ORCHESTRATION', '✅ Script 1 terminé', {
      executionId: executionId
    });
    Logger.log('✅ Script 1 terminé avec succès');
    Logger.log('');

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur Script 1', {
      executionId: executionId,
      error: e.toString(),
      stack: e.stack || ''
    });
    Logger.log('❌ ERREUR Script 1 : ' + e.toString());
    Logger.log('⚠️  Poursuite de la séquence malgré l\'erreur');
    Logger.log('');
  }

  // ═══════════════════════════════════════════════
  // ÉTAPE 2 : Programmation trigger export+purge
  // ═══════════════════════════════════════════════
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('⏰ PROGRAMMATION TRIGGER (dans 3 minutes)');
  Logger.log('═══════════════════════════════════════════════════');

  try {
    ScriptApp.newTrigger('exportEtPurgeRapportClient_')
      .timeBased()
      .after(3 * 60 * 1000)
      .create();

    logStructured_('INFO', 'ORCHESTRATION', '⏰ Trigger programmé', {
      executionId: executionId,
      delayMs: 3 * 60 * 1000,
      nextFunction: 'exportEtPurgeRapportClient_'
    });

    Logger.log('✅ Trigger programmé : exportEtPurgeRapportClient_()');
    Logger.log('   Exécution dans : 3 minutes');
    Logger.log('');

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur programmation trigger', {
      executionId: executionId,
      error: e.toString()
    });
    Logger.log('❌ ERREUR : Impossible de programmer le trigger');
    Logger.log(e.toString());
  }

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';

  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ ORCHESTRATION PHASE 1 TERMINÉE');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 Résumé :');
  Logger.log('   • DELTA calculé   : ' + todoKeys.length + ' client(s)');
  Logger.log('   • Script 1 exécuté: ✅');
  Logger.log('   • Trigger programmé: ✅ (dans 3 min)');
  Logger.log('   • Durée Phase 1   : ' + duration);
  Logger.log('═══════════════════════════════════════════════════');

  logStructured_('INFO', 'ORCHESTRATION', '✅ Phase 1 terminée', {
    executionId: executionId,
    deltaCount: todoKeys.length,
    duration: duration,
    nextPhase: 'exportEtPurgeRapportClient_ (dans 3 min)'
  });
}

/****************************************************
 * ═══════════════════════════════════════════════
 * PHASE 2 : EXPORT + PURGE (Trigger automatique)
 * ═══════════════════════════════════════════════
 ****************************************************/
function exportEtPurgeRapportClient_() {
  var startTime = Date.now();
  var executionId = PropertiesService.getScriptProperties().getProperty('lastExecutionId') || Utilities.getUuid();

  logStructured_('INFO', 'ORCHESTRATION', '🚀 Démarrage Phase 2 (Export+Purge)', { 
    executionId: executionId 
  });
  
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 ORCHESTRATION - PHASE 2 (Export + Purge)');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🆔 Execution ID : ' + executionId);
  Logger.log('');

  // ✅ Nettoyage trigger actuel
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'exportEtPurgeRapportClient_') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  // ═══════════════════════════════════════════════
  // ÉTAPE 2a : Purge sites non débutés (Script 3)
  // ═══════════════════════════════════════════════
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🧹 ÉTAPE 2a : PURGE SITES NON DÉBUTÉS (Script 3)');
  Logger.log('═══════════════════════════════════════════════════');

  try {
    purgeRapportClientNonDebute();
    
    logStructured_('INFO', 'ORCHESTRATION', '✅ Script 3 terminé', {
      executionId: executionId
    });
    Logger.log('✅ Script 3 terminé avec succès');
    Logger.log('');

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur Script 3', {
      executionId: executionId,
      error: e.toString(),
      stack: e.stack || ''
    });
    Logger.log('❌ ERREUR Script 3 : ' + e.toString());
    Logger.log('⚠️  Poursuite de la séquence malgré l\'erreur');
    Logger.log('');
  }

  // ═══════════════════════════════════════════════
  // ÉTAPE 2b : Export données finalisés (Script 2)
  // ═══════════════════════════════════════════════
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📤 ÉTAPE 2b : EXPORT DONNÉES (Script 2)');
  Logger.log('═══════════════════════════════════════════════════');

  try {
    exportRapportClient_AllInOne();
    
    logStructured_('INFO', 'ORCHESTRATION', '✅ Script 2 terminé', {
      executionId: executionId
    });
    Logger.log('✅ Script 2 terminé avec succès');
    Logger.log('');

  } catch(e) {
    logStructured_('ERROR', 'ORCHESTRATION', '❌ Erreur Script 2', {
      executionId: executionId,
      error: e.toString(),
      stack: e.stack || ''
    });
    Logger.log('❌ ERREUR Script 2 : ' + e.toString());
    Logger.log('');
  }

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';

  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ ORCHESTRATION PHASE 2 TERMINÉE');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 Résumé :');
  Logger.log('   • Script 3 exécuté: ✅');
  Logger.log('   • Script 2 exécuté: ✅');
  Logger.log('   • Durée Phase 2   : ' + duration);
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🎉 SÉQUENCE COMPLÈTE TERMINÉE');
  Logger.log('═══════════════════════════════════════════════════');

  logStructured_('INFO', 'ORCHESTRATION', '✅ Phase 2 terminée - Séquence complète', {
    executionId: executionId,
    duration: duration
  });

  // ✅ Nettoyage final
  PropertiesService.getScriptProperties().deleteProperty('lastExecutionId');
}

/****************************************************
 * ═══════════════════════════════════════════════
 * HELPERS ORCHESTRATION
 * ═══════════════════════════════════════════════
 ****************************************************/

/**
 * Nettoie tous les triggers existants de la séquence
 */
function cleanupExistingTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var cleaned = 0;

  for (var t = 0; t < triggers.length; t++) {
    var handler = triggers[t].getHandlerFunction();
    if (handler === 'demarrerSequenceAsynchrone_' ||
        handler === 'exportEtPurgeRapportClient_' ||
        handler === 'exportRapportClient_AllInOne') {
      ScriptApp.deleteTrigger(triggers[t]);
      cleaned++;
    }
  }

  if (cleaned > 0) {
    logStructured_('INFO', 'ORCHESTRATION', '🧹 Triggers nettoyés', { count: cleaned });
    Logger.log('🧹 ' + cleaned + ' trigger(s) existant(s) supprimé(s)');
  }
}

/**
 * Fonction de test manuelle (sans AppSheet)
 */
function testSequenceComplete() {
  Logger.log('🧪 TEST MANUEL - Séquence complète');
  var result = lancerRapportComplet();
  Logger.log('Résultat : ' + JSON.stringify(result));
}

/**
 * Annule tous les triggers en cours
 */
function annulerSequenceEnCours() {
  cleanupExistingTriggers_();
  PropertiesService.getScriptProperties().deleteProperty('lastExecutionId');
  
  logStructured_('WARN', 'ORCHESTRATION', '⚠️ Séquence annulée manuellement', {});
  Logger.log('⚠️ Tous les triggers ont été annulés');
  Logger.log('✅ Vous pouvez relancer lancerRapportComplet()');
}

/**
 * Vérifie le statut de l'exécution en cours
 */
function verifierStatutExecution() {
  var executionId = PropertiesService.getScriptProperties().getProperty('lastExecutionId');
  
  if (!executionId) {
    Logger.log('❌ Aucune exécution en cours');
    return { status: 'idle', message: 'Aucune exécution en cours' };
  }

  var triggers = ScriptApp.getProjectTriggers();
  var activeTriggers = [];
  
  for (var t = 0; t < triggers.length; t++) {
    var handler = triggers[t].getHandlerFunction();
    if (handler === 'demarrerSequenceAsynchrone_' || 
        handler === 'exportEtPurgeRapportClient_') {
      activeTriggers.push(handler);
    }
  }

  Logger.log('📊 Statut exécution :');
  Logger.log('   • Execution ID : ' + executionId);
  Logger.log('   • Triggers actifs : ' + activeTriggers.length);
  Logger.log('   • Fonctions : ' + activeTriggers.join(', '));

  return {
    status: 'running',
    executionId: executionId,
    activeTriggers: activeTriggers
  };
}
