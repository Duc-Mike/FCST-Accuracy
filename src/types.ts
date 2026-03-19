export interface ReferenceData {
  registerWM: Record<string, string>; // Week -> Month
  divisionRegion: Record<string, { division: string; region: string }>; // Sales office/group -> Div/Reg
  category: Record<string, string>; // Division name -> Category
}

export interface MergedRow {
  Code: string;
  'Registration month': string;
  'Plan month': string;
  Division: string;
  Region: string;
  Material: string;
  Customer: string;
  'End customer': string;
  'Sales employee': string;
  'Quantity (PCS)': number;
  'Amount (KRW)': number;
  'Net price': number;
  Category: string;
}
