const TOKEN_KEY = "cpo_attainment_token";
let authToken = null;

export function setToken(token) {
  authToken = token;
  if (token) localStorage.setItem(TOKEN_KEY, token);
  else localStorage.removeItem(TOKEN_KEY);
}
export function getStoredToken() {
  return localStorage.getItem(TOKEN_KEY);
}

async function request(path, opts = {}) {
  const headers = { "Content-Type": "application/json", ...(opts.headers || {}) };
  if (authToken) headers.Authorization = `Bearer ${authToken}`;
  const res = await fetch(`/api${path}`, {
    method: opts.method || "GET",
    headers,
    body: opts.body ? JSON.stringify(opts.body) : undefined,
  });
  let data = {};
  try { data = await res.json(); } catch { /* no body */ }
  if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);
  return data;
}

export const api = {
  register: (username, password, displayName) =>
    request("/auth/register", { method: "POST", body: { username, password, displayName } }),
  login: (username, password) =>
    request("/auth/login", { method: "POST", body: { username, password } }),
  me: () => request("/auth/me"),
  listCourses: () => request("/courses"),
  getCourse: (id) => request(`/courses/${id}`),
  upsertCourse: (courseCode, courseName, data) =>
    request("/courses/upsert", { method: "POST", body: { courseCode, courseName, data } }),
  deleteCourse: (id) => request(`/courses/${id}`, { method: "DELETE" }),
};
