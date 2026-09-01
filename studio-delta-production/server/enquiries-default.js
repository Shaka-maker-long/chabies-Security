function unique(items) {
  const out = [];
  const seen = new Set();
  for (const raw of items) {
    const s = String(raw || "").trim();
    if (!s) continue;
    const key = s.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(s);
  }
  return out;
}

const ENQUIRY_FIELDS = [
  "enquiry_no", "date_enquired", "month_enquired", "enquiry_source", "enquiry_type",
  "client_name", "source", "client_email", "client_number", "province", "category",
  "product", "request", "status", "date_quoted", "quote_no", "comment"
];

const ENQUIRY_DROPDOWN_KEYS = [
  "enquiry_source", "enquiry_type", "source", "status", "province", "category", "product"
];

const DEFAULT_ENQUIRY_DROPDOWNS = {
  enquiry_source: unique([
    "Call", "Email", "Facebook", "Instagram", "No Trace", "Showroom", "Website",
    "Whatsapp", "Delta Request", "Decor & Design"
  ]),
  enquiry_type: unique([
    "Catologue", "Custom", "General", "New Design", "Showroom Appointment"
  ]),
  source: unique([
    "Billboards", "Facebook", "Friends/Family", "Google", "Google Ads", "Instagram",
    "Magazine", "Recurring Client", "No Trace", "Inexss", "Designer Meeting",
    "Decor & Design", "Delta Employee"
  ]),
  status: unique([
    "Costed", "Followed Up", "New", "Ordered", "Quoted", "Re-Cost",
    "Waiting on clients personal details", "Waiting on clients specifictions",
    "Waiting on productions confirmation", "Waiting on Supplier",
    "Not within scope", "Not Interested"
  ]),
  province: unique([
    "Eastern Cape", "Free State", "Gauteng", "KwaZulu-Natal", "Limpopo",
    "Mpumalanga", "North West", "Northern Cape", "Western Cape"
  ]),
  category: unique([
    "Aircon Structure", "Baby Changing Compactum Unit", "Baby Crib", "Back Doors Roofs",
    "Bed", "Bedframe", "Bedside Table", "Bench", "Brackets", "Buffet", "Burthers Block",
    "Cabinet", "Chair", "Coffee Table", "Console", "Couch", "Cupboard",
    "Custom Academy Desks and Modular Drawers", "Custom Lexan Divider", "Custom Steel Box",
    "Custom Two Sides Pin Up Boards", "Custom Upright", "Desk", "Door", "Droppers",
    "Floating Shelf", "Foot Stool", "Frame", "Gate", "Gin Trolley", "Glass", "Headboard",
    "Industrial Bed", "Kitchen Island", "Labour & Sundries", "Locking Mechanism",
    "Magazine Stand", "Mirror", "Ottoman", "Panels", "Patio Roof + Small frames at fire place",
    "Planter", "Plinth", "Rail", "Room Divider", "Server", "Shelf", "Shelf ladders",
    "Side Table", "Sideboard", "Steel Table Legs", "Steel Trolley", "Table", "Tag",
    "TV Unit", "Vanity", "Wardrobe", "Waste Sorting Table", "Wine Pegs", "Wine Racks",
    "Worm Farm", "Partitions", "Delivery"
  ]),
  product: unique([
    "Air Bar Stool", "Air Chair", "Air Nook Chair", "Aircon Structure", "Aliana Dining Table",
    "Alice Corner Couch", "Amalia Side Table", "Amara Bed", "Amelia Arched Cabinet",
    "Anaya Display Cabinet", "Angelica Coffee Table", "Angelica Patio Chair", "Angelica Patio Couch",
    "Angelica Planter Box", "Angelica Plinth", "Arched Coffee Station", "Astra Steel Cabinet",
    "Athena Room Divider", "Atlas Steel Chair", "Back Doors Roofs", "Bathroom Vanity",
    "Bella Side Table", "Blaire Bedside Cabinet", "Blaire Curved Headboard", "Blaire Dresser",
    "Blaire Dressing Table", "Blaire Platform Bed", "Blaire Wardrobe", "Brackets",
    "Brie Drinks Cabinet", "Brielle Bed Frame", "Brielle Bedside cabinet", "Brooke Open Wardrobe",
    "Brooklyn Display Cabinet", "Camilla Display Cabinet", "Camilla Display Cabinet - Brass Plated",
    "Charlene Arched Cabinet", "Charlotte Chair", "Charmaine Overhead Shelf", "Claire Coffee Table",
    "Cupboard with Push to Open Door", "Custom Academy Desks and Modular Drawers",
    "Custom Bathroom Vanity", "Custom Console", "Custom Lexan Divider",
    "Custom Modular Arched Coffee Station", "Custom Modular Bookshelf", "Custom Shoe Cabinet",
    "Custom Steel Box", "Custom Steel Chair", "Custom Storage Sidetable", "Custom Table",
    "Custom Three Tier - Magazine Stand", "Custom TV Unit", "Custom Two Sides Pin Up Boards",
    "Custom Upright", "Daphne Rectangular Mirror", "Delilah Display Cabinet", "Delivery",
    "Diana Cabinet", "Display Cabinet", "Door", "Doors", "Droppers", "Ella Arched Cabinet",
    "Eloise Desk", "Estee Corner Couch", "Estelle Patio Table", "Eve Patio Bench", "Eve Patio Set",
    "Eve Patio Table", "Evelyn Cabinet", "Evie Arched Wardrobe", "Evie Changing Compactum",
    "Evie Crib Cot", "Evie Rocking Chair", "Evie Toy Shelf", "Felicity Floating Shelf - Large",
    "Felicity Floating Shelf - Medium", "Felicity Floating Shelf - Small",
    "Felicity Floating Shelves - Bundle", "Fleur TV Unit", "Floating Bedside Cabinets",
    "Floating Shelf", "Floating Tables", "Foot Stool", "Frame", "Garden Gate", "Giselle Chair",
    "Glass", "Hex Bar Stool", "Hex Dinning Chair", "Hex Nook Chair", "Imani Open Shelf",
    "Jasmine Planter", "Khanya Kitchen Cabinet", "Kitchen Island", "Kyra Nesting Side Tables",
    "Labour & Sundries", "Leonora Bedside Table", "Lila Side Table", "Lindi Console Table",
    "Locking Mechanism", "Lorna Open Shelf", "Lucia Coffee Table", "Luna Console Table",
    "Lyra Steel Chair", "Mabel Wall Mirror", "Madeleine Room Divider", "Margaux Open Wardrobe",
    "Maria Desk", "Matilds TV unit", "Mbali Planter", "Melody Standing Mirror",
    "Michelle Coffee Table", "Mila Dining Table", "Myra Bed Bench", "N/A",
    "Naledi Drinks Cabinet", "Nandi Dining Chair", "Nandi Dining Set", "Nandi Dining Table",
    "Naomi Arched Cabinet", "Noelle Open Shelf", "Nola Side Table", "Nova Bar Stool",
    "Oakland Room Divider", "Oakland Sideboard", "Office Divider", "Olivia Bathroom Cabinet",
    "Ophelia Coffee Table", "Orion Steel Chair", "Ottoman", "Palesa Coffee Table", "Panels",
    "Partitions", "Patio Roof + Small frames at fire place", "Penelope Display Cabinet",
    "Pill shaped Mirror", "Planter Box", "Quinn Patio Table", "Raye Coffee Table", "Rosie Planter",
    "Ruby Drinks Cabinet", "Serena Sideboard", "Shelf", "Sienna Buffet", "Siya Side Table",
    "Solis Steel Chair", "Steel Cabinet", "Steel Gate", "Steel Trolley", "Table", "Tag",
    "Talitha Bookshelf", "Tatiana Bookshelf", "Thalia Display Cabinet", "Thandi Bedside Cabinet",
    "Thandi Display Cabinet", "Themba Bookshelf", "Titan Steel Chair", "Towel Rail", "Tumi TV Unit",
    "Uriah Side Table", "Valerie Cabinet", "Vanessa Floating Shelf - Large",
    "Vanessa Floating Shelf - Medium", "Vanessa Floating Shelf - Small",
    "Vanessa Floating Shelves - Bundle", "Vega Steel Chair", "Vinette Wine Cabinet",
    "Violet Sideboard 2-Door", "Violet Sideboard 3-Door", "Violet Sideboard 4-Door",
    "Vivienne Arched Cabinet", "Vivienne Sideboard 4-Door", "Wall Unit Headboard",
    "Waste Sorting Table", "Wine Pegs", "Wine racks", "Worm Farm", "Zahara Arched Mirror",
    "Zandile TV unit", "Zanele TV Unit", "Zola Console Table", "Zola Side Table",
    "Zuri Sideboard - Oval"
  ])
};

module.exports = {
  unique,
  ENQUIRY_FIELDS,
  ENQUIRY_DROPDOWN_KEYS,
  DEFAULT_ENQUIRY_DROPDOWNS
};
