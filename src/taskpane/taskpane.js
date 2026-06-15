// =========================================================
// ASISTENTE DE COMPRENSIÓN INTELIGENTE - CÓDIGO UNIFICADO
// =========================================================
import "./taskpane.css";
// Variables de estado del complemento
let tipoExplicacion = "sencillo";

// ========== INICIALIZACIÓN Y ESCUCHADORES ==========

Office.onReady(function(info) {
    console.log("Office.js listo - Asistente de Comprensión");
    
    if (info.host === Office.HostType.Word) {
        updateSelectedText();
        verificarDisparoExterno();

        window.addEventListener("storage", function(event) {
            if (event.key === "consultarDesdeMenu" && event.newValue === "true") {
                verificarDisparoExterno();
            }
        });

        // Detectar foco en el panel
        window.addEventListener('focus', function() {
            console.log("Panel enfocado");
            if (localStorage.getItem("consultarDesdeMenu") === "true") {
                verificarDisparoExterno();
            }
        });

        // Actualizar texto seleccionado al hacer clic en el panel
        window.addEventListener('focus', function() {
            console.log("Actualizar selección por foco");
            cargarSeleccionActual();
        });

        const btn = document.getElementById("explainBtn");
        if (btn) {
            btn.onclick = ejecutarConsulta;
        }

        // Botón de papelera - Borrar todo
        const clearBtn = document.getElementById("clearAllBtn");
        if (clearBtn) {
            clearBtn.onclick = borrarTodo;
        }

        // NUEVO ENLACE: Conectar los clicks de las pestañas a la función de cambio visual
        const btnSencillo = document.getElementById("btn-sencillo");
        const btnTecnico = document.getElementById("btn-tecnico");
        const btnEjemplo = document.getElementById("btn-ejemplo");

        if (btnSencillo) btnSencillo.onclick = () => window.setTipo("sencillo");
        if (btnTecnico) btnTecnico.onclick = () => window.setTipo("tecnico");
        if (btnEjemplo) btnEjemplo.onclick = () => window.setTipo("ejemplo");
    }
});

function verificarDisparoExterno() {
    if (localStorage.getItem("consultarDesdeMenu") === "true") {
        localStorage.removeItem("consultarDesdeMenu");
        
        // Borrado directo y forzado
        var resultDiv = document.getElementById("result");
        if (resultDiv) {
            resultDiv.innerHTML = "";
            console.log("Resultado anterior borrado");
        }
        
        setTimeout(function() {
            ejecutarConsulta();
        }, 100);
    }
}

// ========== LÓGICA DE CONSULTA E IA BLINDADA ==========

async function consultarIA(palabra, contexto) {
    const tieneContextoSuficiente = contexto && contexto.trim().length > 20;

    let prompt = "";
    if (tieneContextoSuficiente) {
        prompt = `Explica qué significa "${palabra}" en esta oración: "${contexto}". 
Tipo de explicación: ${tipoExplicacion}

    Instrucciones según tipo:
    - sencillo: explicación básica para estudiantes principiantes
    - tecnico: explicación formal, académica y precisa
    - ejemplo: incluye un ejemplo práctico para facilitar comprensión

Instrucciones:
- Respuesta de máximo 35 palabras.
- No copies la oración del contexto.
- Da una definición clara y, si es relevante, su propósito o una característica clave.
- Usa la estructura: "[Término] es un/una [definición]" o "[Término] se refiere a [definición]".`;
    } else {
        prompt = `Define "${palabra}" de forma general, indicando su propósito o los campos donde se usa. Usa la estructura: "[Término] es un/una [definición]" o "[Término] se refiere a [definición]". Máximo 35 palabras.`;
    }

    try {
        const respuesta = await fetch("https://localhost:3001/ia", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ prompt })
        });
        
        const data = await respuesta.json();
        
        // Si el servidor Node responde con un código de error (Fuera del rango 200-299)
        if (!respuesta.ok) {
            console.error("Error detectado en el backend:", data);
            
            switch (respuesta.status) {
                case 429: // Límite de tokens/solicitudes excedido (Rate Limit)
                    throw new Error(data.details || "Límite de consultas excedido. Por favor, espera un minuto antes de intentar otra consulta.");
                
                case 400: // Selección excesivamente gigante
                    throw new Error("El texto seleccionado es demasiado largo. Intenta seleccionando una frase o palabra más corta.");
                
                case 401: // Error de llaves .env o permisos remotos
                case 403:
                    throw new Error("Error de autenticación: La clave de la API de Gemini no es válida o no está configurada.");
                
                case 500: // Caída interna del código backend
                    throw new Error("El servidor sufrió un error interno al procesar el texto. Inténtalo de nuevo.");
                
                default: // Cualquier otro error imprevisto
                    throw new Error(data.error || "Ocurrió un inconveniente inesperado al comunicarse con la IA.");
            }
        }
        
        return data.respuesta || "Sin respuesta del modelo.";

    } catch (error) {
        console.error("Error capturado en el flujo fetch:", error);
        
        if (error.message && !error.message.includes("Failed to fetch")) {
            return error.message;
        }
        
        return "No se pudo conectar con el asistente. Asegúrate de que el servidor de Node.js esté encendido en la consola.";
    }
}

function ejecutarConsulta() {
    showLoading(true);
    clearResult();

    Word.run(async function(context) {
        try {
            var selection = context.document.getSelection();
            selection.load("text");
            
            var paragraph = selection.paragraphs.getFirst();
            paragraph.load("text");
            
            await context.sync();
            
            var selectedText = selection.text;
            
            // 1. FILTRO: Selección vacía o puros espacios en blanco
            if (!selectedText || selectedText.trim() === "") {
                showResult("No hay texto seleccionado. Selecciona una palabra o frase en tu documento.", "error");
                showLoading(false);
                return;
            }

            // 2. FILTRO: Texto Basura sin vocales (Ej: jjjjkkjjkjkjkkk) o caracteres repetidos (Ej: aaaaaa)
            const limpio = selectedText.trim();
            const tieneVocales = /[aeiouáéíóúü]/i.test(limpio);
            const esLetraRepetida = /^(.)\1+$/.test(limpio);
            
            // 3. FILTRO: Símbolos o Puntuación repetida sin letras (Ej: ......... , ,,,, , ----)
            const tieneLetrasONumeros = /[a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/.test(limpio);

            if ((limpio.length > 5 && !tieneVocales) || esLetraRepetida || !tieneLetrasONumeros) {
                showResult("⚠ La frase o texto seleccionado no es correcto para analizar.", "error");
                showLoading(false);
                return; // Frena la ejecución, impidiendo el envío innecesario al servidor
            }

            updateSelectedTextDisplay(selectedText);
            
            var contextoCompleto = paragraph.text || selectedText;
            
            // Realizar la consulta al flujo blindado si pasó todos los filtros anteriores
            var explicacion = await consultarIA(selectedText, contextoCompleto);
            
            // DETECTAR TIPO DE RESPUESTA VISUAL:
            if (
                explicacion.includes("Límite de consultas") || 
                explicacion.includes("demasiado largo") || 
                explicacion.includes("autenticación") || 
                explicacion.includes("No se pudo conectar") ||
                explicacion.includes("error interno")
            ) {
                showResult("⚠ " + explicacion, "error");
            } else {
                showResult(explicacion, "success");
            }
            
            showLoading(false);

        } catch (error) {
            console.error(error);
            showResult("Error al procesar la solicitud: " + error.message, "error");
            showLoading(false);
        }
    }).catch(function(error) {
        console.error(error);
        showResult("Error al leer el documento. Asegúrate de tener un archivo abierto en Word.", "error");
        showLoading(false);
    });
}

// ========== FUNCIONES DE INTERFAZ (UI) ==========

function updateSelectedText() {
    Word.run(function(context) {
        var selection = context.document.getSelection();
        selection.load("text");
        return context.sync().then(function() {
            updateSelectedTextDisplay(selection.text);
        });
    }).catch(function() {});
}

function updateSelectedTextDisplay(text) {
    var el = document.getElementById("selectedDisplay");
    if (el) {
        if (text && text.trim()) {
            var displayText = text.length > 100 ? text.substring(0, 100) + "..." : text;
            el.innerHTML = "\"" + displayText + "\"";
            el.classList.remove("empty");
        } else {
            el.innerHTML = "Ninguno";
            el.classList.add("empty");
        }
    }
}

function showResult(msg, type) {
    var resultDiv = document.getElementById("result");
    var bg = "#e8f0fe";
    var border = "#667eea";
    
    if (type === "success") {
        bg = "#e6f4ea";
        border = "#34a853";
    } else if (type === "error") {
        bg = "#fce8e6";
        border = "#ea4335";
    }
    
    var mensajeLimpio = msg.replace(/\*\*/g, "").replace(/\*/g, "");
    
    if (resultDiv) {
        resultDiv.innerHTML =
        "<div class='result-box' style='background:" + bg + "; border-left-color:" + border + ";'>" +
            mensajeLimpio.replace(/\n/g, "<br>") +
        "</div>";
    }

    const btnSpan = document.querySelector("#explainBtn span");
    if (btnSpan) {
        btnSpan.textContent = "Generar otra explicación";
    }
}

function clearResult() {
    var resultDiv = document.getElementById("result");
    if (resultDiv) {
        resultDiv.innerHTML = "";
        console.log("clearResult ejecutado");
    }
}

function showLoading(show) {
    var loader = document.getElementById("loading");
    if (loader) {
        if (show) {
            loader.classList.remove("hidden");
        } else {
            loader.classList.add("hidden");
        }
    }
}

function borrarTodo() {
    const resultDiv = document.getElementById("result");
    if (resultDiv) {
        resultDiv.innerHTML = "";
    }
    
    const selectedDisplay = document.getElementById("selectedDisplay");
    if (selectedDisplay) {
        selectedDisplay.innerHTML = "Ninguno";
        selectedDisplay.classList.add("empty");
    }
    
    const loader = document.getElementById("loading");
    if (loader) {
        loader.classList.add("hidden");
    }
    
    const btnSpan = document.querySelector("#explainBtn span");
    if (btnSpan) {
        btnSpan.textContent = "Generar explicación"; // <--- Ajustado para que haga juego con el HTML
    }
}

function cargarSeleccionActual() {
    Word.run(function(context) {
        var selection = context.document.getSelection();
        selection.load("text");
        return context.sync().then(function() {
            updateSelectedTextDisplay(selection.text);
        });
    }).catch(function(error) {
        console.log("No se pudo actualizar selección:", error);
    });
}

window.setTipo = function(tipo) {
    tipoExplicacion = tipo;

    const slider = document.getElementById("pill-slider");
    if (!slider) return;

    const positions = {
        sencillo: "0%",
        tecnico: "100%",
        ejemplo: "200%"
    };

    slider.style.transform = `translateX(${positions[tipo]})`;
    slider.className = "pill-slider " + tipo;

    document.querySelectorAll(".pill-container button").forEach(btn => {
        btn.classList.remove("active");
    });

    const activeBtn = document.getElementById("btn-" + tipo);
    if (activeBtn) activeBtn.classList.add("active");

    slider.animate([
        { transform: `translateX(${positions[tipo]}) scale(0.95)` },
        { transform: `translateX(${positions[tipo]}) scale(1)` }
    ], {
        duration: 200,
        easing: "cubic-bezier(0.4, 0, 0.2, 1)"
    });
};

window.addEventListener("load", () => {
    document.querySelector(".pill-container")?.animate([
        { transform: "scale(0.98)" },
        { transform: "scale(1)" }
    ], {
        duration: 120,
        easing: "ease-out"
    });
});