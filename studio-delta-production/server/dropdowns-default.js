function unique(items) {
  const skip = new Set([
    "type", "catergory", "category", "product", "variation", "doors",
    "powder coating", "province", "source"
  ]);
  const out = [];
  const seen = new Set();
  for (const raw of items) {
    const s = String(raw || "").trim();
    if (!s || skip.has(s.toLowerCase())) continue;
    const key = s.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(s);
  }
  return out;
}

const DROPDOWN_KEYS = [
  "type",
  "category",
  "product",
  "variation",
  "doors",
  "powder_coating",
  "province",
  "source"
];

const DROPDOWN_LABELS = {
  type: "TYPE",
  category: "CATERGORY",
  product: "PRODUCT",
  variation: "VARIATION",
  doors: "DOORS",
  powder_coating: "POWDER COATING",
  province: "PROVINCE",
  source: "SOURCE"
};

const DEFAULT_DROPDOWNS = {
  type: unique(["Standard", "Custom", "New Design"]),
  category: unique([
    "Aircon Structure", "Baby Changing Compactum Unit", "Baby Crib", "Back Doors Roofs",
    "Bedframe", "Bedside Table", "Bench", "Brackets", "Buffet", "Burthers Block", "Cabinet",
    "Cage", "Chair", "Coffee Table", "Console", "Couch", "Cupboard",
    "Custom Academy Desks and Modular Drawers", "Custom Lexan Divider", "Custom Steel Box",
    "Custom Two Sides Pin Up Boards", "Desk", "Door", "Dresser", "Droppers", "Floating Shelf",
    "Frame", "Gate", "Gin Trolley", "Glass", "Hanging Frame", "Headboard", "Industrial Bed",
    "Kitchen Island", "Labour and Sundries", "Magazine Stand", "Mirror", "Ottoman", "Panels",
    "Patio Roof + Small frames at fire place", "Planter", "Plinth", "Room Divider", "Server",
    "Shelf", "Shelf ladders", "Side Table", "Sideboard", "Space Saver Dinning Set", "Steel Frame",
    "Steel Table Legs", "Table", "TV Unit", "Vanity", "Wardrobe", "Waste Sorting Table",
    "Wine Pegs", "Wine Racks", "Worm Farm", "Custom Upright", "Custom Tags", "Custom Steel"
  ]),
  product: unique([
    "Air Bar Stool", "Air Chair", "Air Nook Chair", "Aircon Structure", "Aliana Dining Table",
    "Alice Corner Couch", "Amalia Side Table", "Amara Bed", "Amelia Arched Cabinet",
    "Anaya Display Cabinet", "Angelica Coffee Table", "Angelica Patio Chair", "Angelica Patio Couch",
    "Angelica Planter Box", "Angelica Plinth", "Arched Coffee Station", "Astra Steel Cabinet",
    "Atlas Steel Chair", "Back Doors Roofs", "Bathroom Vanity", "Bella Side Table",
    "Blaire Bedside Cabinet", "Blaire Curved Headboard", "Blaire Dresser", "Blaire Dressing Table",
    "Blaire Platform Bed", "Blaire Wardrobe", "Brackets", "Brie Drinks Cabinet", "Brielle Bed Frame",
    "Brielle Bedside cabinet", "Brooke Open Wardrobe", "Brooklyn Display Cabinet", "Cage",
    "Camilla Display Cabinet", "Camilla Display Cabinet - Brass Plated", "Charlene Arched Cabinet",
    "Charlotte Chair", "Charmaine Overhead Shelf", "Claire Coffee Table", "Concrete Table",
    "Cupboard with Push to Open Door", "Custom Academy Desks and Modular Drawers",
    "Custom Bathroom Vanity", "Custom Bench Table", "Custom Lexan Divider",
    "Custom Modular Arched Coffee Station", "Custom Modular Bookshelf", "Custom Shoe Cabinet",
    "Custom Steel Box", "Custom Steel Chair", "Custom Steel Frame", "Custom Storage Sidetable",
    "Custom Table", "Custom Tags", "Custom Three Tier - Magazine Stand", "Custom TV Unit",
    "Custom Two Sides Pin Up Boards", "Custom Upright", "Daphne Rectangular Mirror",
    "Delilah Display Cabinet", "Diana Cabinet", "Door", "Doors", "Droppers", "Ella Arched Cabinet",
    "Eloise Desk", "Estee Corner Couch", "Estelle Patio Table", "Eve Patio Bench", "Eve Patio Set",
    "Eve Patio Table", "Evelyn Cabinet", "Evie Arched Wardrobe", "Evie Changing Compactum",
    "Evie Crib Cot", "Evie Rocking Chair", "Evie Toy Shelf", "Felicity Floating Shelf - Large",
    "Felicity Floating Shelf - Medium", "Felicity Floating Shelf - Small",
    "Felicity Floating Shelves - Bundle", "Fleur TV Unit", "Floating Bedside Cabinets",
    "Floating Shelf", "Floating Tables", "Frame", "Garden Gate", "Giselle Chair", "Glass",
    "Hanging Frame", "Hex Bar Stool", "Hex Dinning Chair", "Hex Nook Chair", "Imani Open Shelf",
    "Jasmine Planter", "Khanya Kitchen Cabinet", "Kitchen Island", "Kyra Nesting Side Tables",
    "Labour and Sundries", "Leonora Bedside Table", "Lila Side Table", "Lindi Console Table",
    "Lorna Open Shelf", "Lucia Coffee Table", "Lucille Leather Chair", "Luna Console Table",
    "Lyra Steel Chair", "Madeleine Room Divider", "Margaux Open Wardrobe", "Maria Desk",
    "Marisol Coffee Table", "Matilds TV unit", "Mbali Planter", "Melody Standing Mirror",
    "Michelle Coffee Table", "Mila Dining Table", "Mirror", "Myra Bed Bench", "N/A",
    "Naledi Drinks Cabinet", "Nandi Dining Chair", "Nandi Dining Set", "Nandi Dining Table",
    "Naomi Arched Cabinet", "Noelle Open Shelf", "Nola Side Table", "Nova Bar Stool",
    "Oakland Room Divider", "Oakland Sideboard", "Office Divider", "Olivia Bathroom Cabinet",
    "Ophelia Coffee Table", "Orion Steel Chair", "Ottoman", "Palesa Coffee Table", "Panels",
    "Patio Roof + Small frames at fire place", "Penelope Display Cabinet", "Pill shaped Mirror",
    "Planter Box", "Pole", "Quinn Patio Table", "Raye Coffee Table", "Rosie Planter",
    "Ruby Drinks Cabinet", "Serena Sideboard", "Sienna Buffet", "Siya Side Table",
    "Solis Steel Chair", "Steel Cabinet", "Steel Gate", "Swivel Drinks Cabinet", "Table Frame",
    "Talitha Bookshelf", "Tatiana Bookshelf", "Thalia Display Cabinet", "Thandi Bedside Cabinet",
    "Thandi Display Cabinet", "Themba Bookshelf", "Titan Steel Chair", "Tumi TV Unit",
    "Uriah Side Table", "Valerie Cabinet", "Vanessa Floating Shelf - Large",
    "Vanessa Floating Shelf - Medium", "Vanessa Floating Shelf - Small",
    "Vanessa Floating Shelves - Bundle", "Vega Steel Chair", "Vinette Wine Cabinet",
    "Violet Sideboard 2-Door", "Violet Sideboard 3-Door", "Violet Sideboard 4-Door",
    "Vivienne Arched Cabinet", "Vivienne Sideboard 4-Door", "Wall Unit Headboard",
    "Waste Sorting Table", "Wine Pegs", "Wine racks", "Worm Farm", "Zahara Arched Mirror",
    "Zandile TV unit", "Zanele TV Unit", "Zola Console Table", "Zola Side Table",
    "Zuri Sideboard - Oval", "Custom Steel"
  ]),
  variation: unique([
    "Amara Bed - 3/4 Extra Length: 1100mm W x 2020mm L x 1200mm H",
    "Amara Bed - Double Extra Length: 1400mm W x 2020mm L x 1200mm H",
    "Amara Bed - Double: 1400mm W x 1900mm L x 1200mm H",
    "Amara Bed - King Extra Lenght:1890mm W x 2020mm L x 1200mm H",
    "Amara Bed - Queen Extra Length: 1580mm w X 2020mm L x 1200mm H",
    "Amara Bed- 3/4: 1100mm W x 1900mm L x 1200mm H",
    "Amara Bed- King: 1890mm W x 1900mm L x 1200mm H",
    "Amara Bed- Queen: 1580mm W x 1900mm L x 1200mm H",
    "Angelica Planter - Large: 1200mm W x 500mm D x 670mm H",
    "Angelica Planter - Short: 700mm W x 500mm D x 670mm H",
    "Angelica Plinth - Medium: 350mm W x 350mm D x 800mm H",
    "Angelica Plinth - Short:: 300mm W x 300mm D x 600mm H",
    "Angelica Plinth - Tall: 400mm W x 400mm D x 1000mm H",
    "As per website",
    "Ash wood drawers",
    "Blaire Platform Bed - 3/4 : 1120mm W x 1900mm L x 300mm H",
    "Blaire Platform Bed - 3/4 Extra Length: 1120mm W x 2020mm L x 300mm H",
    "Blaire Platform Bed - Double Extra Length: 1420mm W x 2020mm L x 300mm H",
    "Blaire Platform Bed - Double: 1420mm W x 1900mm L x 300mm H",
    "Blaire Platform Bed - King Extra Length: 1870mm W x 2020mm L x 300mm H",
    "Blaire Platform - Bed: King: 1870mm W x 1900mm L x 300mm H",
    "Blaire Platform Bed: Queen Extra Length: 1560mm W x 2020mm L x 300mm H",
    "Blaire Platform Bed: Queen: 1560mm W x 1900mm L x 300mm H",
    "Blaire Platform Bed: Single: 960mm W x 1900mm L x 300mm W",
    "Clear glass",
    "Felicity Floating Shelf - Large: 900mm W x 250mm D x 300mm H",
    "Felicity Floating Shelf - Medium: 600mm W x 250mm D x 300mm H",
    "Felicity Floating Shelf - Small: 450mm W x 250mm D x 300mm H",
    "Jasmine Planter - Medium: 300mm W x 300mm D x 1000mm H",
    "Jasmine Planter - Short: 300mm W x 300mm D x 700mm H",
    "Jasmine Planter: Tall: 300mm W x 300mm D x 1200mm H",
    "N/A",
    "Porcelain Terrazzo Top",
    "Reeded Glass",
    "Refer to description",
    "Rosie Planter - Medium: 350mm Di x 1000mm H",
    "Rosie Planter - Short: 350mm Di x 700mm H",
    "Rosie Planter - Tall: 350mm Di x 1300mm H",
    "Top & bottom shelves glass, Middles shelves: Glass",
    "Top & bottom shelves glass, Middles shelves: Steel",
    "Top & bottom shelves glass, Middles shelves: Wood",
    "Top & bottom shelves steel, Middles shelves: Glass",
    "Top & bottom shelves steel, Middles shelves: Steel",
    "Top & bottom shelves steel, Middles shelves: Wood",
    "Top & bottom shelves wood, Middles shelves: Glass",
    "Top & bottom shelves wood, Middles shelves: Steel",
    "Top & bottom shelves wood, Middles shelves: Wood",
    "Top and bottom glass",
    "Top and bottom shelves steel: Bottom Wine Rack",
    "Top and bottom shelves steel: Top shelves wood: Bottom Wine Rack",
    "Top and bottom steel",
    "Top glass, bottom steel",
    "Top glass, bottom wood.",
    "Top steel, bottom glass",
    "Vanessa Floating Shelf - Large: 900mm W x 200mm D x 300mm H",
    "Vanessa Floating Shelf - Medium: 600mm W x 200mm D x 300mm H",
    "Vanessa Floating Shelf - Small: 300mm W x 200mm D x 300mm H",
    "Wood top, middle and bottom steel",
    "Internal Wine Shelving",
    "Tile Top and Bottom"
  ]),
  doors: unique([
    "As per website", "Clear glass", "Clear Glass and Steel Doors", "CNC Perforated Steel",
    "Engraved MDF", "Mirror", "N/A", "No Doors", "Ocean Vue", "Perforated Steel", "Reeded Glass",
    "Reeded Glass and Steel Doors", "Refer to description", "Skinny Reeded - Grey", "Steel",
    "Wood", "Woven Steel Doors", "Black Skinny Reeded Glass"
  ]),
  powder_coating: unique([
    "As per website", "Does not require powder coating", "Ferrograin Black", "Ferrograin White",
    "Ferrograin Wine Red (Maroon)", "FerroTex Black", "Ivory", "New Gold", "Refer to description",
    "Sea Foam", "Smooth Matt Black", "Sepia Brown"
  ]),
  province: unique([
    "Eastern Cape", "Free State", "Gauteng", "KwaZulu-Natal", "Limpopo", "Mpumalanga",
    "North West", "Northern Cape", "Western Cape"
  ]),
  source: unique([
    "Billboards", "Friends/Family", "Google", "Google Ads", "Magazine", "No Trace",
    "Recurring Client", "Social Media"
  ])
};

module.exports = { unique, DROPDOWN_KEYS, DROPDOWN_LABELS, DEFAULT_DROPDOWNS };
