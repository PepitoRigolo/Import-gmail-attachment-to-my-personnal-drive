
/**
 * Récupère et sauvegarde les pièces jointes des mails ayant le libellé passé en paramètre.  
 * Archive et marque comme lu les mails traités.  
 * @param {LibelleConfig} libelleConfig
 * @param {string[]} messageIdHistory - ids des messages à ignorer
 * @returns
 */
function saveAttachmentsFromLibelleConfig(libelleConfig, messageIdHistory)
{
    let threads = GmailApp.search(`in:inbox has:attachment label:${libelleConfig.libelle}`);
    console.info(`${threads.length} ${threads.length > 1 ? 'threads trouvés' : 'thread trouvé'} pour le libellé '${libelleConfig.libelle}'`);
    
    if (threads.length == 0) return;

    let parentFolder = getFolderFromPath(libelleConfig.folderPath);

    threads.forEach(thread => {
        try
        {
            thread.getMessages().forEach(message => {
                if (messageIdHistory.includes(message.getId())) return;
                message.getAttachments().forEach(attachment => {
                    console.info(`Enregistrement de la pièces jointe '${attachment.getName()}' du mail '${message.getSubject()}' - '${message.getId()}' envoyé par '${message.getFrom()}' ayant le libellé '${libelleConfig.libelle}'`);
                    parentFolder.createFile(attachment.copyBlob());
                    appendHistory(message.getId(), libelleConfig.libelle, message.getFrom(), attachment.getName(), libelleConfig.folderPath);
                });
                message.markRead();
            });
            thread.moveToArchive();
        }
        catch (e)
        {
            console.warn(`Erreur lors du traitement du thread '${thread.getFirstMessageSubject()}' : ${e}`);
        }
    });
}