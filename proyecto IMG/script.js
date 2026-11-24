
"use strict";

const $ = id => document.getElementById(id);

const LIBRARY_KEY = "lumina_songs_v3";
const LIVE_KEY = "lumina_projection_v3";

const isAudience =
    new URLSearchParams(window.location.search)
        .get("view") === "audience";

let songs = [];
let selectedId = null;
let selectedSlide = 0;
let projectorWindow = null;
let notificationTimer;

const channel = typeof BroadcastChannel !== "undefined"
    ? new BroadcastChannel("lumina_projection")
    : null;

// ========================================
// ALMACENAMIENTO DE CANCIONES
// ========================================

function readStorage(key, fallback) {
    try {
        return JSON.parse(localStorage.getItem(key)) ?? fallback;
    } catch {
        return fallback;
    }
}

function loadSongs() {
    songs = readStorage(LIBRARY_KEY, []);
}

function saveSongs() {
    try {
        localStorage.setItem(
            LIBRARY_KEY,
            JSON.stringify(songs)
        );
        return true;
    } catch {
        notify("No se pudo guardar. Almacenamiento lleno.");
        return false;
    }
}

function getSelectedSong() {
    return songs.find(song => song.id === selectedId);
}

function getLive() {
    return readStorage(LIVE_KEY, null);
}

function createId() {
    return globalThis.crypto?.randomUUID?.()
        || Date.now() + "-" + Math.random();
}

function notify(message) {
    const notification = $("notification");

    notification.textContent = message;
    notification.classList.add("visible");

    clearTimeout(notificationTimer);

    notificationTimer = setTimeout(() => {
        notification.classList.remove("visible");
    }, 3000);
}

// ========================================
// BIBLIOTECA DE CANCIONES
// ========================================

function renderLibrary() {
    const library = $("library");
    const search = $("search").value.toLowerCase();
    const live = getLive();

    library.replaceChildren();

    const filtered = songs.filter(song =>
        song.title.toLowerCase().includes(search)
    );

    if (!filtered.length) {
        const message = document.createElement("p");
        message.className = "muted";
        message.textContent = "No hay canciones disponibles.";
        library.appendChild(message);
    }

    filtered.forEach(song => {

        const button = document.createElement("button");

        button.className = "song-item";

        if (song.id === selectedId) {
            button.classList.add("active");
        }

        const icon = document.createElement("span");
        icon.className = "song-icon";
        icon.textContent = "♫";

        const info = document.createElement("span");
        info.className = "song-info";

        const title = document.createElement("strong");
        title.textContent = song.title;

        const count = document.createElement("small");
        count.textContent =
            song.slides.length + " diapositivas";

        info.append(title, count);

        if (live && live.id === song.id) {
            const status = document.createElement("small");
            status.className = "on-air";
            status.textContent = "● EN PROYECCIÓN";
            info.appendChild(status);
        }

        button.append(icon, info);

        button.addEventListener("click", () => {
            selectSong(song.id);
        });

        library.appendChild(button);
    });
}

// ========================================
// SELECCIÓN AUTOMÁTICA DE POWERPOINT
// ========================================

function selectSong(id) {
    selectedId = id;
    selectedSlide = 0;

    renderLibrary();
    renderEditor();

    // Si el modo privado está desactivado,
    // la canción se proyecta automáticamente.

    if (!$("privateMode").checked) {
        publishSong(id, 0);
    }
}

// ========================================
// CREAR NUEVA CANCIÓN
// ========================================

function newSong() {
    const song = {
        id: createId(),
        title: "Nueva canción",
        slides: ["Escribe aquí la primera estrofa"]
    };

    songs.push(song);

    saveSongs();

    // Evita mostrar una canción sin terminar.
    $("privateMode").checked = true;

    updateMode();
    selectSong(song.id);

    notify("Nueva canción creada en privado.");
}

// ========================================
// IMPORTAR ARCHIVOS POWERPOINT
// ========================================

async function importPowerPoint(file) {

    if (!file.name.toLowerCase().endsWith(".pptx")) {
        notify("El archivo debe estar en formato PPTX.");
        return;
    }

    if (typeof JSZip === "undefined") {
        notify("Comprueba tu conexión a internet.");
        return;
    }

    const zip = await JSZip.loadAsync(file);

    const parser = new DOMParser();

    const drawingNS =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    const presentationNS =
        "http://schemas.openxmlformats.org/presentationml/2006/main";

    const relationNS =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    let slideFiles = [];

    // Leer el orden real de las diapositivas.

    const presentationFile =
        zip.file("ppt/presentation.xml");

    const relationsFile =
        zip.file("ppt/_rels/presentation.xml.rels");

    if (presentationFile && relationsFile) {

        const presentation = parser.parseFromString(
            await presentationFile.async("string"),
            "application/xml"
        );

        const relations = parser.parseFromString(
            await relationsFile.async("string"),
            "application/xml"
        );

        const relationMap = new Map();

        Array.from(
            relations.getElementsByTagNameNS("*", "Relationship")
        ).forEach(relation => {

            relationMap.set(
                relation.getAttribute("Id"),
                relation.getAttribute("Target")
            );
        });

        const slideIds = Array.from(
            presentation.getElementsByTagNameNS(
                presentationNS,
                "sldId"
            )
        );

        slideFiles = slideIds.map(slide => {

            const id = slide.getAttributeNS(
                relationNS,
                "id"
            );

            const target = relationMap.get(id);

            if (!target) return null;

            const path = new URL(
                target,
                "https://ppt.local/ppt/"
            ).pathname.slice(1);

            return path;

        }).filter(path => path && zip.file(path));
    }

    // Alternativa si no se encuentra el orden.

    if (!slideFiles.length) {

        slideFiles = Object.keys(zip.files)
            .filter(path =>
                /^ppt\/slides\/slide\d+\.xml$/.test(path)
            )
            .sort((a, b) => {

                const numberA = Number(
                    a.match(/slide(\d+)/)[1]
                );

                const numberB = Number(
                    b.match(/slide(\d+)/)[1]
                );

                return numberA - numberB;
            });
    }

    if (!slideFiles.length) {
        throw new Error("No se encontraron diapositivas.");
    }

    const slides = [];

    for (const path of slideFiles) {

        const xml = await zip.file(path).async("string");

        const documentXML = parser.parseFromString(
            xml,
            "application/xml"
        );

        const paragraphs = Array.from(
            documentXML.getElementsByTagNameNS(
                drawingNS,
                "p"
            )
        );

        const lines = paragraphs.map(paragraph => {

            let line = "";

            Array.from(paragraph.children).forEach(node => {

                if (node.localName === "br") {
                    line += "\n";
                } else {
                    line += Array.from(
                        node.getElementsByTagNameNS(
                            drawingNS,
                            "t"
                        )
                    ).map(text => text.textContent).join("");
                }
            });

            return line;
        });

        slides.push(lines.join("\n").trim());
    }

    const song = {
        id: createId(),
        title: file.name.replace(/\.pptx$/i, ""),
        slides: slides
    };

    songs.push(song);

    saveSongs();

    // Importar no interrumpe la proyección actual.

    $("privateMode").checked = true;

    updateMode();
    selectSong(song.id);

    notify("Canción importada: " + song.title);
}

// ========================================
// IMPORTACIÓN MÚLTIPLE
// ========================================

async function importFiles(event) {

    const files = Array.from(event.target.files);

    for (const file of files) {
        try {
            await importPowerPoint(file);
        } catch (error) {
            notify("Error al importar: " + file.name);
            console.error(error);
        }
    }

    event.target.value = "";
}

// ========================================
// EDITOR PRIVADO
// ========================================

function renderEditor() {

    const song = getSelectedSong();

    $("editor").hidden = !song;

    if (!song) return;

    selectedSlide = Math.max(
        0,
        Math.min(selectedSlide, song.slides.length - 1)
    );

    $("songTitle").value = song.title;

    $("lyrics").value =
        song.slides[selectedSlide] || "";

    $("privatePreview").textContent =
        song.slides[selectedSlide] || "";

    renderSlideList();
}

function renderSlideList() {

    const song = getSelectedSong();

    if (!song) return;

    const list = $("slideList");

    list.replaceChildren();

    song.slides.forEach((lyrics, index) => {

        const button = document.createElement("button");

        button.className = "slide-item";

        if (index === selectedSlide) {
            button.classList.add("selected");
        }

        const title = document.createElement("strong");

        title.textContent = "Estrofa " + (index + 1);

        const preview = document.createElement("span");

        preview.textContent =
            lyrics.slice(0, 85) || "Diapositiva vacía";

        button.append(title, preview);

        button.onclick = () => {

            selectedSlide = index;

            renderEditor();

            // En modo directo, seleccionar una estrofa
            // la muestra inmediatamente en el proyector.

            if (!$("privateMode").checked) {
                publishSong(song.id, index);
            }
        };

        list.appendChild(button);
    });
}

// ========================================
// CORREGIR LETRAS SIN MOSTRAR LOS CAMBIOS
// ========================================

function editLyrics() {

    const song = getSelectedSong();

    if (!song) return;

    // Al editar se activa automáticamente
    // el modo de preparación privada.

    $("privateMode").checked = true;

    updateMode();

    song.slides[selectedSlide] = $("lyrics").value;

    $("privatePreview").textContent =
        song.slides[selectedSlide];

    saveSongs();

    renderSlideList();
}

function editTitle() {

    const song = getSelectedSong();

    if (!song) return;

    $("privateMode").checked = true;

    updateMode();

    song.title = $("songTitle").value;

    saveSongs();

    renderLibrary();
}

// ========================================
// AÑADIR Y ELIMINAR ESTROFAS
// ========================================

function addSlide() {

    const song = getSelectedSong();

    if (!song) return;

    $("privateMode").checked = true;

    updateMode();

    song.slides.splice(selectedSlide + 1, 0, "");

    selectedSlide++;

    saveSongs();

    renderEditor();

    $("lyrics").focus();
}

function removeSlide() {

    const song = getSelectedSong();

    if (!song) return;

    if (song.slides.length <= 1) {
        notify("Debe quedar al menos una estrofa.");
        return;
    }

    if (!confirm("¿Eliminar esta estrofa?")) return;

    $("privateMode").checked = true;

    updateMode();

    song.slides.splice(selectedSlide, 1);

    selectedSlide = Math.max(0, selectedSlide - 1);

    saveSongs();

    renderEditor();
}

function removeSong() {

    const song = getSelectedSong();

    if (!song) return;

    if (!confirm("¿Eliminar " + song.title + "?")) return;

    songs = songs.filter(item => item.id !== song.id);

    selectedId = null;

    saveSongs();

    renderLibrary();
    renderEditor();
    renderLive();

    notify("Canción eliminada de la biblioteca.");
}

// ========================================
// MODO PRIVADO Y MODO DIRECTO
// ========================================

function updateMode() {

    const privateMode = $("privateMode").checked;

    $("modeInfo").textContent = privateMode
        ? "Modo privado: el público no verá tus cambios."
        : "Modo directo: las canciones se proyectan automáticamente.";
}

function changeMode() {

    updateMode();

    // Al salir del modo privado, la canción
    // corregida se proyecta automáticamente.

    if (!$("privateMode").checked && selectedId) {

        publishSong(selectedId, selectedSlide);

        notify("Canción corregida en proyección.");

    } else {

        notify("Modo de preparación privada activado.");
    }
}

// ========================================
// PUBLICAR EN EL PROYECTOR
// ========================================

function publishSong(id, index = 0) {

    const song = songs.find(item => item.id === id);

    if (!song) return;

    // Se crea una copia independiente de las letras.
    // Editar el original no modifica esta copia.

    const live = {
        id: song.id,
        title: song.title,
        slides: [...song.slides],
        index: Math.max(
            0,
            Math.min(index, song.slides.length - 1)
        ),
        black: false
    };

    sendLive(live);
}

function sendLive(live) {

    try {

        localStorage.setItem(
            LIVE_KEY,
            JSON.stringify(live)
        );

    } catch {

        notify("No se pudo actualizar la proyección.");
        return;
    }

    if (channel) {
        channel.postMessage("update");
    }

    if (isAudience) {
        renderAudience();
    } else {
        renderLive();
        renderLibrary();
    }
}

// ========================================
// VISTA EN VIVO DEL ADMINISTRADOR
// ========================================

function renderLive() {

    const live = getLive();

    $("prevSlide").disabled =
        !live || live.index <= 0;

    $("nextSlide").disabled =
        !live || live.index >= live.slides.length - 1;

    $("blackBtn").disabled = !live;

    $("prevSong").disabled = songs.length === 0;
    $("nextSong").disabled = songs.length === 0;

    if (!live) {

        $("liveName").textContent =
            "Ninguna canción activa";

        $("liveCounter").textContent =
            "Sin proyección";

        $("livePreview").textContent =
            "Esperando una canción...";

        return;
    }

    $("liveName").textContent = live.title;

    $("liveCounter").textContent =
        "Estrofa " + (live.index + 1)
        + " de " + live.slides.length;

    $("livePreview").textContent = live.black
        ? ""
        : live.slides[live.index];

    $("blackBtn").textContent = live.black
        ? "□ Mostrar letras"
        : "■ Pantalla negra";

    fitText($("livePreview"));
}

// ========================================
// CAMBIAR ESTROFAS
// ========================================

function changeSlide(direction) {

    const live = getLive();

    if (!live) return;

    const nextIndex = live.index + direction;

    if (
        nextIndex < 0 ||
        nextIndex >= live.slides.length
    ) {
        return;
    }

    live.index = nextIndex;

    sendLive(live);

    if (
        !isAudience &&
        !$("privateMode").checked &&
        selectedId === live.id
    ) {

        selectedSlide = live.index;

        renderEditor();
    }
}

// ========================================
// CAMBIAR POWERPOINT AUTOMÁTICAMENTE
// ========================================

function changeSong(direction) {

    if (!songs.length) return;

    const live = getLive();

    const currentIndex = songs.findIndex(
        song => song.id === live?.id
    );

    let nextIndex;

    if (currentIndex === -1) {

        nextIndex = direction > 0
            ? 0
            : songs.length - 1;

    } else {

        nextIndex = currentIndex + direction;
    }

    if (nextIndex >= songs.length) {

        notify("Has llegado a la última canción.");
        return;
    }

    if (nextIndex < 0) {

        notify("Estás en la primera canción.");
        return;
    }

    const nextSong = songs[nextIndex];

    // Cambia directamente la proyección,
    // sin solicitar una confirmación adicional.

    publishSong(nextSong.id, 0);

    // Si estás corrigiendo otra canción,
    // conserva el editor privado abierto.

    if (!$("privateMode").checked) {

        selectedId = nextSong.id;
        selectedSlide = 0;

        renderEditor();
    }

    renderLibrary();
}

// ========================================
// PANTALLA NEGRA
// ========================================

function toggleBlack() {

    const live = getLive();

    if (!live) return;

    live.black = !live.black;

    sendLive(live);
}

// ========================================
// ABRIR VENTANA DEL PROYECTOR
// ========================================

function openProjector() {

    const url = new URL(window.location.href);

    url.searchParams.set("view", "audience");

    projectorWindow = window.open(
        url.href,
        "LuminaProjector",
        "width=1280,height=720"
    );

    if (!projectorWindow) {

        notify("Permite las ventanas emergentes.");
        return;
    }

    notify("Mueve la ventana al proyector y presiona F.");
}

// ========================================
// AJUSTAR TAMAÑO DE LAS LETRAS
// ========================================

function fitText(element) {

    const parent = element.parentElement;

    if (!parent) return;

    element.style.fontSize = "";

    let size = parseFloat(
        getComputedStyle(element).fontSize
    );

    const availableHeight = isAudience
        ? window.innerHeight * 0.9
        : element.clientHeight;

    while (
        element.scrollHeight > availableHeight &&
        size > 12
    ) {

        size -= 1;

        element.style.fontSize = size + "px";
    }
}

// ========================================
// PANTALLA PÚBLICA
// ========================================

function renderAudience() {

    const live = getLive();

    const text = $("audienceText");

    if (!live || live.black) {

        text.textContent = "";
        return;
    }

    // El público solamente verá la letra.

    text.textContent =
        live.slides[live.index] || "";

    fitText(text);
}

function fullscreen() {

    if (!document.fullscreenElement) {

        document.documentElement
            .requestFullscreen?.()
            .catch(console.error);

    } else {

        document.exitFullscreen?.()
            .catch(console.error);
    }
}

// ========================================
// ATAJOS DE TECLADO
// ========================================

function keyboardControls(event) {

    const tag = document.activeElement?.tagName;

    if (["INPUT", "TEXTAREA"].includes(tag)) return;

    if (
        event.ctrlKey &&
        event.key === "ArrowRight" &&
        !isAudience
    ) {

        event.preventDefault();
        changeSong(1);

    } else if (
        event.ctrlKey &&
        event.key === "ArrowLeft" &&
        !isAudience
    ) {

        event.preventDefault();
        changeSong(-1);

    } else if (
        event.key === "ArrowRight" ||
        event.key === "PageDown"
    ) {

        event.preventDefault();
        changeSlide(1);

    } else if (
        event.key === "ArrowLeft" ||
        event.key === "PageUp"
    ) {

        event.preventDefault();
        changeSlide(-1);

    } else if (event.key.toLowerCase() === "b") {

        toggleBlack();

    } else if (
        event.key.toLowerCase() === "f" &&
        isAudience
    ) {

        fullscreen();
    }
}

// ========================================
// SINCRONIZACIÓN ENTRE LAS DOS PANTALLAS
// ========================================

function syncProjection() {

    if (isAudience) {

        renderAudience();

    } else {

        renderLive();
        renderLibrary();
    }
}

window.addEventListener("storage", event => {

    if (event.key === LIVE_KEY) {
        syncProjection();
    }
});

if (channel) {

    channel.addEventListener("message", () => {
        syncProjection();
    });
}

document.addEventListener(
    "keydown",
    keyboardControls
);

// ========================================
// INICIALIZACIÓN DEL PROGRAMA
// ========================================

function init() {

    if (isAudience) {

        $("admin").hidden = true;
        $("audience").hidden = false;

        document.title = "LÚMINA | Pantalla pública";

        renderAudience();

        window.addEventListener(
            "resize",
            renderAudience
        );

        $("audience").ondblclick = fullscreen;

        return;
    }

    loadSongs();

    $("search").oninput = renderLibrary;

    $("importBtn").onclick = () => {
        $("files").click();
    };

    $("files").onchange = importFiles;

    $("newBtn").onclick = newSong;

    $("openProjector").onclick = openProjector;

    $("privateMode").onchange = changeMode;

    $("songTitle").oninput = editTitle;

    $("lyrics").oninput = editLyrics;

    $("addSlide").onclick = addSlide;

    $("removeSlide").onclick = removeSlide;

    $("removeSong").onclick = removeSong;

    $("prevSlide").onclick = () => changeSlide(-1);

    $("nextSlide").onclick = () => changeSlide(1);

    $("prevSong").onclick = () => changeSong(-1);

    $("nextSong").onclick = () => changeSong(1);

    $("blackBtn").onclick = toggleBlack;

    renderLibrary();
    renderEditor();
    renderLive();
    updateMode();
}

init();
