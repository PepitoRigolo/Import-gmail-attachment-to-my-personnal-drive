
/**
 * Fonction principale du script
 */
function execution() {
    const libelleConfigs = getLibelleConfigs();
    const messageIdHistory = getHistory();

    libelleConfigs.forEach(libelleConfig => {
        saveAttachmentsFromLibelleConfig(libelleConfig, messageIdHistory);
    });
}
