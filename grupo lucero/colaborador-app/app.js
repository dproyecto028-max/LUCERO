const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxM1CXu082VlcKkOelVleIcrRvcGKsuJ-aMdhc2qW6Tf3IH3CmmKkAg3MRGN2AucsWaAw/exec";

const state = {
  usuario: null,
  jornadaActiva: false,
  cronometro: null,
  inicioMillis: null,
  gpsInicio: null,
  gpsFin: null,
  timerInterval: null
};

const loginScreen = document.getElementById("loginScreen");
const dashboardScreen = document.getElementById("dashboardScreen");
const btnLogout = document.getElementById("btnLogout");
const timerEl = document.getElementById("cronometro");
const usuarioNombreEl = document.getElementById("usuarioNombre");
const estadoJornadaEl = document.getElementById("estadoJornada");
const lugarTrabajoInput = document.getElementById("lugarTrabajoInput");
const novedadesInput = document.getElementById("novedadesInput");
const comprobanteEl = document.getElementById("comprobanteJornada");

function setEstadoJornada(texto, tipo = "neutral") {
  estadoJornadaEl.textContent = texto;
  estadoJornadaEl.className = `status-pill ${tipo}`;
}

function formatTiempo(ms) {
  const totalSegundos = Math.max(0, Math.floor(ms / 1000));
  const horas = String(Math.floor(totalSegundos / 3600)).padStart(2, "0");
  const minutos = String(Math.floor((totalSegundos % 3600) / 60)).padStart(2, "0");
  const segundos = String(totalSegundos % 60).padStart(2, "0");
  return `${horas}:${minutos}:${segundos}`;
}

function startTimer() {
  if (state.timerInterval) {
    clearInterval(state.timerInterval);
  }

  state.inicioMillis = Date.now();
  state.timerInterval = setInterval(() => {
    const elapsed = Date.now() - state.inicioMillis;
    timerEl.textContent = formatTiempo(elapsed);
  }, 1000);
}

function stopTimer() {
  if (state.timerInterval) {
    clearInterval(state.timerInterval);
    state.timerInterval = null;
  }
}

function updateUI() {
  if (!state.usuario) {
    loginScreen.classList.remove("hidden");
    dashboardScreen.classList.add("hidden");
    btnLogout.classList.add("hidden");
    return;
  }

  loginScreen.classList.add("hidden");
  dashboardScreen.classList.remove("hidden");
  btnLogout.classList.remove("hidden");
  usuarioNombreEl.textContent = state.usuario.nombre || "Colaborador";

  if (state.jornadaActiva) {
    setEstadoJornada("Jornada en curso", "active");
    if (state.inicioMillis) {
      timerEl.textContent = formatTiempo(Date.now() - state.inicioMillis);
    }
  } else {
    setEstadoJornada("Jornada no iniciada", "neutral");
    timerEl.textContent = "00:00:00";
  }
}

function guardarSesion() {
  localStorage.setItem("cs_colab_user", JSON.stringify(state.usuario));
}

function cargarSesion() {
  const saved = localStorage.getItem("cs_colab_user");
  if (!saved) return;

  try {
    state.usuario = JSON.parse(saved);
  } catch (error) {
    state.usuario = null;
  }

  updateUI();
}

function logout() {
  state.usuario = null;
  state.jornadaActiva = false;
  state.inicioMillis = null;
  state.gpsInicio = null;
  state.gpsFin = null;
  stopTimer();
  localStorage.removeItem("cs_colab_user");
  document.getElementById("loginForm").reset();
  document.getElementById("novedadesInput").value = "";
  document.getElementById("lugarTrabajoInput").value = "";
  comprobanteEl.innerHTML = "Aún no se generó un comprobante.";
  updateUI();
}

function requestLocation() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error("Tu navegador no soporta GPS."));
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (pos) => {
        resolve({
          lat: pos.coords.latitude,
          lng: pos.coords.longitude,
          precision: pos.coords.accuracy
        });
      },
      (err) => reject(err),
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
    );
  });
}

async function loginColaborador(event) {
  event.preventDefault();

  const dni = document.getElementById("loginDni").value.trim();
  const password = document.getElementById("loginPass").value.trim();

  if (!dni || !password) {
    alert("Completá DNI y contraseña.");
    return;
  }

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        action: "login",
        dni,
        password
      })
    });

    const data = await res.json();

    if (!data || data.status !== "success" || data.type !== "colaborador") {
      alert(data.message || "Credenciales inválidas o el usuario no es colaborador.");
      return;
    }

    state.usuario = {
      nombre: data.name,
      email: data.email,
      tipo: data.type
    };

    guardarSesion();
    updateUI();
  } catch (error) {
    alert("No se pudo iniciar sesión. Verificá la conexión o la URL.");
  }
}

async function iniciarJornada() {
  if (!state.usuario) {
    alert("Debés iniciar sesión primero.");
    return;
  }

  const lugarTrabajo = lugarTrabajoInput.value.trim();
  if (!lugarTrabajo) {
    alert("Ingresá el lugar de trabajo antes de iniciar la jornada.");
    return;
  }

  try {
    const ubicacion = await requestLocation();
    state.gpsInicio = ubicacion;

    const payload = {
      action: "marcar_jornada",
      tipo: "inicio",
      email: state.usuario.email,
      nombre: state.usuario.nombre,
      lugarTrabajo,
      lat: ubicacion.lat,
      lng: ubicacion.lng,
      precision: ubicacion.precision
    };

    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    if (!data || data.status !== "success") {
      alert(data.message || "No se pudo iniciar la jornada.");
      return;
    }

    state.jornadaActiva = true;
    startTimer();
    setEstadoJornada("Jornada en curso", "active");
  } catch (error) {
    alert("Necesitás habilitar el acceso a la ubicación GPS para iniciar la jornada.");
  }
}

async function finalizarJornada() {
  if (!state.usuario || !state.jornadaActiva) {
    alert("Primero debés iniciar la jornada.");
    return;
  }

  try {
    const ubicacion = await requestLocation();
    state.gpsFin = ubicacion;

    const payload = {
      action: "marcar_jornada",
      tipo: "fin",
      email: state.usuario.email,
      nombre: state.usuario.nombre,
      lugarTrabajo: lugarTrabajoInput.value.trim(),
      lat: ubicacion.lat,
      lng: ubicacion.lng,
      precision: ubicacion.precision,
      novedades: novedadesInput.value.trim() || "Sin novedades"
    };

    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    if (!data || data.status !== "success") {
      alert(data.message || "No se pudo finalizar la jornada.");
      return;
    }

    stopTimer();
    state.jornadaActiva = false;
    setEstadoJornada("Jornada finalizada", "active");

    const total = timerEl.textContent || "00:00:00";
    comprobanteEl.innerHTML = `
      <strong>Resumen de jornada</strong><br>
      Colaborador: ${state.usuario.nombre}<br>
      Lugar: ${lugarTrabajoInput.value.trim()}<br>
      Duración: ${total}<br>
      GPS inicio: ${state.gpsInicio ? `${state.gpsInicio.lat}, ${state.gpsInicio.lng}` : "No disponible"}<br>
      GPS cierre: ${state.gpsFin ? `${state.gpsFin.lat}, ${state.gpsFin.lng}` : "No disponible"}
    `;
  } catch (error) {
    alert("No se pudo cerrar la jornada. Verificá los permisos del GPS.");
  }
}

async function enviarReporte() {
  if (!state.usuario) {
    alert("Debés iniciar sesión antes de reportar.");
    return;
  }

  const descripcion = novedadesInput.value.trim();
  if (!descripcion) {
    alert("Escribí una novedad antes de enviar el reporte.");
    return;
  }

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        action: "guardar_reporte",
        nombreColaborador: state.usuario.nombre,
        motivo: "Novedad laboral",
        descripcion,
        lugarTrabajo: lugarTrabajoInput.value.trim() || "No especificado"
      })
    });

    const data = await res.json();
    alert(data.message || "Reporte enviado correctamente.");
    novedadesInput.value = "";
  } catch (error) {
    alert("No se pudo enviar el reporte.");
  }
}

document.getElementById("loginForm").addEventListener("submit", loginColaborador);
document.getElementById("btnIniciarJornada").addEventListener("click", iniciarJornada);
document.getElementById("btnFinalizarJornada").addEventListener("click", finalizarJornada);
document.getElementById("btnReporte").addEventListener("click", enviarReporte);
document.getElementById("btnLogout").addEventListener("click", logout);

cargarSesion();
updateUI();
