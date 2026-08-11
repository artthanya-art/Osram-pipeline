import { supabase } from "./supabaseClient";

// Maps the app's camelCase field names to Postgres snake_case columns.
const ENTRY_FIELD_MAP = {
  salesId: "sales_id",
  salesName: "sales_name",
  zone: "zone",
  customerId: "customer_id",
  customerName: "customer_name",
  customerType: "customer_type",
  customerSegment: "customer_segment",
  productType: "product_type",
  itemCode: "item_code",
  itemDescription: "item_description",
  qty: "qty",
  uom: "uom",
  price: "price",
  competitorName: "competitor_name",
  competitorPrice: "competitor_price",
  projectCloseYear: "project_close_year",
  projectCloseMonth: "project_close_month",
  quotationNumber: "quotation_number",
  deliveryMethod: "delivery_method",
  kpiRegister: "kpi_register",
  actionPlanMonth: "action_plan_month",
  progress: "progress",
  deliverInMonths: "deliver_in_months",
  deliveryStartYear: "delivery_start_year",
  deliveryStartMonth: "delivery_start_month",
  startWorkingMonth: "start_working_month",
  visitDate: "visit_date",
  grade: "grade",
};
const NUMERIC_ENTRY_FIELDS = ["qty", "price", "competitorPrice"];

function entryToDb(e) {
  const row = {};
  Object.entries(ENTRY_FIELD_MAP).forEach(([js, col]) => {
    let v = e[js];
    if (NUMERIC_ENTRY_FIELDS.includes(js)) {
      row[col] = v === "" || v === undefined || v === null ? null : Number(v);
    } else {
      row[col] = v === undefined ? null : v;
    }
  });
  return row;
}
function entryFromDb(row) {
  const e = { id: row.id, createdAt: row.created_at, updatedAt: row.updated_at };
  Object.entries(ENTRY_FIELD_MAP).forEach(([js, col]) => {
    e[js] = row[col] ?? "";
  });
  return e;
}

function userFromDb(row) {
  return {
    salesId: row.sales_id,
    password: row.password,
    name: row.name,
    zone: row.zone,
    role: row.role,
    status: row.status,
    createdAt: row.created_at,
  };
}
function userToDb(u) {
  return {
    sales_id: u.salesId,
    password: u.password,
    name: u.name,
    zone: u.zone,
    role: u.role,
    status: u.status,
  };
}

// ---------------- users ----------------
export async function fetchUsers() {
  const { data, error } = await supabase.from("users").select("*").order("created_at", { ascending: true });
  if (error) throw error;
  return data.map(userFromDb);
}
export async function insertUser(user) {
  const { error } = await supabase.from("users").insert(userToDb(user));
  if (error) throw error;
}
export async function updateUserFields(salesId, fields) {
  const dbFields = {};
  if (fields.status) dbFields.status = fields.status;
  if (fields.role) dbFields.role = fields.role;
  if (fields.password) dbFields.password = fields.password;
  const { error } = await supabase.from("users").update(dbFields).eq("sales_id", salesId);
  if (error) throw error;
}
export async function replaceAllUsers(userList) {
  const { error: delError } = await supabase.from("users").delete().neq("sales_id", "__none__");
  if (delError) throw delError;
  if (userList.length) {
    const { error } = await supabase.from("users").insert(userList.map(userToDb));
    if (error) throw error;
  }
}

// ---------------- zones ----------------
export async function fetchZones() {
  const { data, error } = await supabase.from("zones").select("name").order("name", { ascending: true });
  if (error) throw error;
  return data.map((r) => r.name);
}
export async function insertZone(name) {
  const { error } = await supabase.from("zones").insert({ name });
  if (error && error.code !== "23505") throw error; // ignore duplicate-key conflicts
}

// ---------------- pipeline entries ----------------
export async function fetchEntries() {
  const { data, error } = await supabase.from("pipeline_entries").select("*").order("created_at", { ascending: false });
  if (error) throw error;
  return data.map(entryFromDb);
}
export async function insertEntry(entry) {
  const { data, error } = await supabase.from("pipeline_entries").insert(entryToDb(entry)).select().single();
  if (error) throw error;
  return entryFromDb(data);
}
export async function updateEntry(id, fields) {
  const { error } = await supabase
    .from("pipeline_entries")
    .update({ ...entryToDb(fields), updated_at: new Date().toISOString() })
    .eq("id", id);
  if (error) throw error;
}
export async function deleteEntryById(id) {
  const { error } = await supabase.from("pipeline_entries").delete().eq("id", id);
  if (error) throw error;
}
export async function bulkInsertEntries(entryList) {
  const { data, error } = await supabase.from("pipeline_entries").insert(entryList.map(entryToDb)).select();
  if (error) throw error;
  return data.map(entryFromDb);
}
export async function replaceAllEntries(entryList) {
  const { error: delError } = await supabase.from("pipeline_entries").delete().neq("id", "00000000-0000-0000-0000-000000000000");
  if (delError) throw delError;
  if (entryList.length) await bulkInsertEntries(entryList);
}
