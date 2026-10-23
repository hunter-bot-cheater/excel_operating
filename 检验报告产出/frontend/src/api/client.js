const baseURL = import.meta.env.VITE_API_BASE_URL || "/api/v1";

export async function generateInspectionReports(formData) {
  const response = await fetch(`${baseURL}/inspection-reports/generate`, { method: "POST", body: formData });
  return response.json();
}

export async function getTask(taskId) {
  const response = await fetch(`${baseURL}/inspection-reports/tasks/${taskId}`);
  return response.json();
}
