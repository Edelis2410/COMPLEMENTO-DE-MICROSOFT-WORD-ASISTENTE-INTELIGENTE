/*
 * ASISTENTE DE COMPRENSIÓN INTELIGENTE - GESTOR DE COMANDOS OCULTOS
 */

/* global Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    console.log("Hilo de comandos ocultos inicializado correctamente.");
  }
});

/**
 * Función que se ejecuta instantáneamente al hacer clic derecho en Word
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  console.log("Disparo desde el menú contextual detectado.");

  // 1. Activamos el interruptor en el almacenamiento local para que taskpane.js sepa que debe buscar texto
  localStorage.setItem("consultarDesdeMenu", "true");
  
  // 2. Forzamos la apertura programática del panel lateral (así Word lo inicializa a la fuerza si estaba dormido)
  Office.addon.showAsTaskpane().then(function() {
      // 3. Avisamos a Word que la tarea invisible terminó con éxito para que libere el cursor
      event.completed(); 
  }).catch(function(error) {
      console.error("Error al desplegar el panel desde el comando oculto:", error);
      event.completed();
  });
}

// Registrar la función de manera global para que el manifest.xml la localice bajo el nombre "action"
Office.actions.associate("action", action);