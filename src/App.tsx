import { ChangeEvent, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';

type RawRow = {
  region: string;
  service: string;
  revenueTY?: number;
  revenueLY?: number;
  inspectionsTY?: number;
  inspectionsLY?: number;
  totalInspectionsTY?: number;
  totalInspectionsLY?: number;
  channel?: string;
  carParkSizeTY?: number;
  carParkSizeLY?: number;
};

type RegionSummary = {
  name: string;
  services: Record<string, RawRow[]>;
  totalRevenueTY: number;
  totalRevenueLY: number;
  totalInspectionsTY: number;
  totalInspectionsLY: number;
  totalMobileInspectionsTY: number;
  totalMobileInspectionsLY: number;
  mobileB2BInspectionsTY: number;
  nonMandatoryRevenueTY: number;
  carParkSizeTY?: number;
  carParkSizeLY?: number;
};

type Recommendation = {
  type: 'warning' | 'tip';
  message: string;
};

const SAMPLE_PLACEHOLDER = `Region, Service, Revenue_TY, Revenue_LY, Inspections_TY, Inspections_LY, TotalInspections_TY, CarParkSize_TY, CarParkSize_LY, Channel
Abu Dhabi, Standard Inspection, 5200000, 4800000, 12000, 11000, 15000, , , 
Abu Dhabi, Premium Inspection, 1800000, 1500000, 3000, 2600, 15000, , , 
Abu Dhabi, Mobile Inspection B2B, 900000, 750000, 800, 700, 3000, , , B2B
Abu Dhabi, Mobile Inspection B2C, 400000, 350000, 500, 450, 3000, , , B2C
Abu Dhabi, Cross Selling, 600000, 520000, 2200, 1900, 15000, , , 
Abu Dhabi, Non-Mandatory, 1100000, 950000, , , , , , 
Northern Emirates, Standard Inspection, 2100000, 2300000, 7000, 7800, 9000, 180000, 175000, 
`;

const formatAmount = (value?: number) => {
  if (value == null || Number.isNaN(value)) return 'N/A';
  return `AED ${value.toLocaleString('en-US', { maximumFractionDigits: 0 })}`;
};

const formatPercent = (value?: number) => {
  if (value == null || Number.isNaN(value)) return 'N/A';
  return `${value.toFixed(1)}%`;
};

const parseNumber = (value: string | undefined) => {
  if (!value) return undefined;
  const clean = value
    .replace(/AED/gi, '')
    .replace(/%/g, '')
    .replace(/[^0-9.+-]/g, '')
    .trim();
  if (clean === '') return undefined;
  const num = Number(clean);
  return Number.isNaN(num) ? undefined : num;
};

const normalizeKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/_+/g, '_')
    .trim();

const fieldMap: Record<string, keyof RawRow> = {
  region: 'region',
  area: 'region',
  service: 'service',
  service_type: 'service',
  revenue_ty: 'revenueTY',
  revenue_this_year: 'revenueTY',
  revenue_ly: 'revenueLY',
  revenue_last_year: 'revenueLY',
  inspections_ty: 'inspectionsTY',
  inspections_t_y: 'inspectionsTY',
  inspections_ly: 'inspectionsLY',
  inspections_l_y: 'inspectionsLY',
  totalinspections_ty: 'totalInspectionsTY',
  total_inspections_ty: 'totalInspectionsTY',
  totalinspections_ly: 'totalInspectionsLY',
  total_inspections_ly: 'totalInspectionsLY',
  channel: 'channel',
  carparksize_ty: 'carParkSizeTY',
  car_park_size_ty: 'carParkSizeTY',
  carparksize_ly: 'carParkSizeLY',
  car_park_size_ly: 'carParkSizeLY',
};

const parseCSVLine = (line: string, delimiter: string) => {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
};

const guessDelimiter = (line: string) => {
  if (line.includes(',')) return ',';
  if (line.includes('\t')) return '\t';
  if (line.includes(';')) return ';';
  return ',';
};

const parseRows = (text: string): RawRow[] => {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (!lines.length) return [];

  const delimiter = guessDelimiter(lines[0]);
  const headerFields = parseCSVLine(lines[0], delimiter).map(normalizeKey);
  const hasHeader = headerFields.some((field) => Object.keys(fieldMap).includes(field));

  const rows: RawRow[] = [];
  for (let i = hasHeader ? 1 : 0; i < lines.length; i += 1) {
    const rowFields = parseCSVLine(lines[i], delimiter);
    const row: RawRow = {
      region: '',
      service: '',
    };

    rowFields.forEach((value, index) => {
      const key = hasHeader ? fieldMap[headerFields[index]] : undefined;
      if (!key) return;
      if (key === 'region' || key === 'service' || key === 'channel') {
        (row[key] as string) = value;
      } else {
        row[key] = parseNumber(value);
      }
    });

    if (row.region || row.service) {
      rows.push(row);
    }
  }

  return rows;
};

const getTextFromSheet = (sheet: XLSX.WorkSheet) => {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false }) as unknown[][];
  return rows
    .map((row) => row.map((cell) => (cell == null ? '' : String(cell))).join(', '))
    .join('\n');
};

const parseXlsxFile = async (file: File) => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return getTextFromSheet(sheet);
};

const buildSummary = (rows: RawRow[]) => {
  const regions: Record<string, RegionSummary> = {};

  const normalizeService = (serviceValue: string) => {
    const lower = serviceValue.toLowerCase();
    if (lower.includes('premium')) return 'Premium Inspection';
    if (lower.includes('standard')) return 'Standard Inspection';
    if (lower.includes('mobile')) return 'Mobile Inspection';
    if (lower.includes('cross')) return 'Cross Selling';
    if (lower.includes('non-mandatory') || lower.includes('non mandatory')) return 'Non-Mandatory Services';
    return serviceValue;
  };

  rows.forEach((row) => {
    const regionKey = row.region.trim() || 'Unknown';
    const name = regionKey;
    if (!regions[regionKey]) {
      regions[regionKey] = {
        name,
        services: {},
        totalRevenueTY: 0,
        totalRevenueLY: 0,
        totalInspectionsTY: 0,
        totalInspectionsLY: 0,
        totalMobileInspectionsTY: 0,
        totalMobileInspectionsLY: 0,
        mobileB2BInspectionsTY: 0,
        nonMandatoryRevenueTY: 0,
      };
    }

    const normalizedService = normalizeService(row.service || '');
    const summary = regions[regionKey];
    summary.services[normalizedService] ??= [];
    summary.services[normalizedService].push(row);

    const revenueTY = row.revenueTY ?? 0;
    const revenueLY = row.revenueLY ?? 0;
    const inspectionsTY = row.inspectionsTY ?? 0;
    const inspectionsLY = row.inspectionsLY ?? 0;

    summary.totalRevenueTY += revenueTY;
    summary.totalRevenueLY += revenueLY;
    summary.totalInspectionsTY += inspectionsTY;
    summary.totalInspectionsLY += inspectionsLY;

    if (normalizedService === 'Mobile Inspection') {
      summary.totalMobileInspectionsTY += inspectionsTY;
      summary.totalMobileInspectionsLY += inspectionsLY;
      if ((row.channel || '').toLowerCase().includes('b2b')) {
        summary.mobileB2BInspectionsTY += inspectionsTY;
      }
    }

    if (normalizedService === 'Non-Mandatory Services') {
      summary.nonMandatoryRevenueTY += revenueTY;
    }

    if (row.carParkSizeTY) summary.carParkSizeTY = row.carParkSizeTY;
    if (row.carParkSizeLY) summary.carParkSizeLY = row.carParkSizeLY;
  });

  return Object.values(regions);
};

const getAggValues = (rows: RawRow[]) => {
  const totalTY = rows.reduce((sum, row) => sum + (row.revenueTY ?? 0), 0);
  const totalLY = rows.reduce((sum, row) => sum + (row.revenueLY ?? 0), 0);
  const totalInspectionsTY = rows.reduce((sum, row) => sum + (row.inspectionsTY ?? 0), 0);
  return { totalTY, totalLY, totalInspectionsTY };
};

const createChartData = (region: RegionSummary) => {
  const services = ['Standard Inspection', 'Premium Inspection', 'Mobile Inspection', 'Cross Selling', 'Non-Mandatory Services'];
  return services.map((service) => {
    const rows = region.services[service] ?? [];
    const { totalTY, totalLY } = getAggValues(rows);
    return {
      name: service,
      thisYear: totalTY,
      lastYear: totalLY,
    };
  });
};

const computeGrowth = (current?: number, prior?: number) => {
  if (current == null || prior == null || prior === 0) return undefined;
  return ((current - prior) / prior) * 100;
};

const numericFieldKeys = [
  'revenueTY',
  'revenueLY',
  'inspectionsTY',
  'inspectionsLY',
  'totalInspectionsTY',
  'totalInspectionsLY',
  'carParkSizeTY',
  'carParkSizeLY',
] as const;

type NumericRawField = (typeof numericFieldKeys)[number];

const sumValues = (rows: RawRow[], field: NumericRawField) =>
  rows.reduce((sum, row) => sum + ((row[field] as number) ?? 0), 0);

const computeRegionMetrics = (region: RegionSummary) => {
  const standardRows = region.services['Standard Inspection'] ?? [];
  const premiumRows = region.services['Premium Inspection'] ?? [];
  const mobileRows = region.services['Mobile Inspection'] ?? [];
  const crossRows = region.services['Cross Selling'] ?? [];
  const nonMandatoryRows = region.services['Non-Mandatory Services'] ?? [];

  const standardRevenueTY = sumValues(standardRows, 'revenueTY');
  const standardRevenueLY = sumValues(standardRows, 'revenueLY');
  const standardInspectionsTY = sumValues(standardRows, 'inspectionsTY');
  const standardInspectionsLY = sumValues(standardRows, 'inspectionsLY');

  const premiumRevenueTY = sumValues(premiumRows, 'revenueTY');
  const premiumRevenueLY = sumValues(premiumRows, 'revenueLY');
  const premiumInspectionsTY = sumValues(premiumRows, 'inspectionsTY');

  const mobileRevenueTY = sumValues(mobileRows, 'revenueTY');
  const mobileRevenueLY = sumValues(mobileRows, 'revenueLY');
  const mobileInspectionsTY = sumValues(mobileRows, 'inspectionsTY');

  const mobileB2BRows = mobileRows.filter((row) => (row.channel || '').toLowerCase().includes('b2b'));
  const mobileB2BInspectionsTY = sumValues(mobileB2BRows, 'inspectionsTY');

  const crossRevenueTY = sumValues(crossRows, 'revenueTY');
  const crossRevenueLY = sumValues(crossRows, 'revenueLY');
  const crossInspectionsTY = sumValues(crossRows, 'inspectionsTY');

  const nonMandatoryRevenueTY = sumValues(nonMandatoryRows, 'revenueTY');
  const nonMandatoryRevenueLY = sumValues(nonMandatoryRows, 'revenueLY');

  const standardCarParkTY = region.carParkSizeTY;
  const standardCarParkLY = region.carParkSizeLY;

  const premiumConversion = premiumInspectionsTY && region.totalInspectionsTY ? (premiumInspectionsTY / region.totalInspectionsTY) * 100 : undefined;
  const mobileB2BConversion = mobileB2BInspectionsTY && mobileInspectionsTY ? (mobileB2BInspectionsTY / mobileInspectionsTY) * 100 : undefined;
  const crossConversion = crossInspectionsTY && region.totalInspectionsTY ? (crossInspectionsTY / region.totalInspectionsTY) * 100 : undefined;
  const nonMandatoryShare = region.totalRevenueTY ? (nonMandatoryRevenueTY / region.totalRevenueTY) * 100 : undefined;

  const standardCarParkShareTY = standardCarParkTY && standardInspectionsTY ? (standardInspectionsTY / standardCarParkTY) * 100 : undefined;
  const standardCarParkShareLY = standardCarParkLY && standardInspectionsLY ? (standardInspectionsLY / standardCarParkLY) * 100 : undefined;

  return {
    standard: {
      revenueTY: standardRevenueTY,
      revenueLY: standardRevenueLY,
      growth: computeGrowth(standardRevenueTY, standardRevenueLY),
      carParkShareTY: standardCarParkShareTY,
      carParkShareLY: standardCarParkShareLY,
    },
    premium: {
      revenueTY: premiumRevenueTY,
      revenueLY: premiumRevenueLY,
      growth: computeGrowth(premiumRevenueTY, premiumRevenueLY),
      conversion: premiumConversion,
    },
    mobile: {
      revenueTY: mobileRevenueTY,
      revenueLY: mobileRevenueLY,
      growth: computeGrowth(mobileRevenueTY, mobileRevenueLY),
      b2bConversion: mobileB2BConversion,
    },
    cross: {
      revenueTY: crossRevenueTY,
      revenueLY: crossRevenueLY,
      growth: computeGrowth(crossRevenueTY, crossRevenueLY),
      conversion: crossConversion,
    },
    nonMandatory: {
      revenueTY: nonMandatoryRevenueTY,
      revenueLY: nonMandatoryRevenueLY,
      growth: computeGrowth(nonMandatoryRevenueTY, nonMandatoryRevenueLY),
      share: nonMandatoryShare,
    },
  };
};

const buildRecommendations = (regions: RegionSummary[]) => {
  const recommendations: Recommendation[] = [];
  regions.forEach((region) => {
    const metrics = computeRegionMetrics(region);
    const regionLabel = region.name;

    const checks = [
      { label: 'Standard Inspection', growth: metrics.standard.growth },
      { label: 'Premium Inspection', growth: metrics.premium.growth },
      { label: 'Mobile Inspection', growth: metrics.mobile.growth },
      { label: 'Cross Selling', growth: metrics.cross.growth },
      { label: 'Non-Mandatory Services', growth: metrics.nonMandatory.growth },
    ];

    checks.forEach((item) => {
      if (item.growth != null && item.growth < 0) {
        recommendations.push({
          type: 'warning',
          message: `⚠️ ${item.label} revenue declined ${Math.abs(item.growth).toFixed(1)}% in ${regionLabel} — review pricing or demand drivers`,
        });
      }
    });

    if (metrics.premium.conversion != null && metrics.premium.conversion < 10) {
      recommendations.push({
        type: 'tip',
        message: `💡 Premium Inspection conversion is low in ${regionLabel} — consider bundling or promotional campaigns`,
      });
    }

    if (metrics.mobile.b2bConversion != null && metrics.mobile.b2bConversion < 15) {
      recommendations.push({
        type: 'tip',
        message: `💡 B2B mobile channel is underperforming in ${regionLabel} — consider corporate outreach or fleet partnerships`,
      });
    }

    if (metrics.nonMandatory.share != null && metrics.nonMandatory.share < 20) {
      recommendations.push({
        type: 'tip',
        message: `💡 Non-mandatory services are underpenetrated in ${regionLabel} — explore upselling at point of inspection`,
      });
    }

    if (
      region.name.toLowerCase().includes('northern') &&
      metrics.standard.carParkShareTY != null &&
      metrics.standard.carParkShareLY != null &&
      metrics.standard.carParkShareTY < metrics.standard.carParkShareLY
    ) {
      recommendations.push({
        type: 'warning',
        message: `⚠️ Car park share declined in Northern Emirates — assess competitor activity or capacity constraints`,
      });
    }
  });
  return recommendations;
};

const buildExportCsv = (regions: RegionSummary[]) => {
  const lines = ['Region,Service,Revenue_TY,Revenue_LY,Inspections_TY,Inspections_LY,Channel,CarParkSize_TY,CarParkSize_LY'];
  regions.forEach((region) => {
    Object.entries(region.services).forEach(([service, rows]) => {
      rows.forEach((row) => {
        const line = [
          region.name,
          `"${service}"`,
          row.revenueTY ?? '',
          row.revenueLY ?? '',
          row.inspectionsTY ?? '',
          row.inspectionsLY ?? '',
          row.channel ?? '',
          row.carParkSizeTY ?? '',
          row.carParkSizeLY ?? '',
        ].join(',');
        lines.push(line);
      });
    });
  });
  return lines.join('\n');
};

const Card = ({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) => (
  <div className="rounded-3xl border border-slate-700 bg-slate-900/90 p-5 shadow-xl shadow-slate-950/20">
    <div className="mb-4 text-sm uppercase tracking-[0.24em] text-slate-400">{title}</div>
    <div className="space-y-3">{children}</div>
  </div>
);

const KPIValue = ({
  label,
  current,
  previous,
  growth,
  extra,
}: {
  label: string;
  current: string;
  previous: string;
  growth?: number;
  extra?: string;
}) => {
  const positive = growth != null && growth >= 0;
  return (
    <div className="rounded-2xl bg-slate-950/80 p-4 border border-slate-800">
      <div className="flex items-center justify-between text-sm text-slate-400">{label}</div>
      <div className="mt-3 flex flex-col gap-2">
        <div className="text-2xl font-semibold text-white">{current}</div>
        <div className="text-sm text-slate-400">Last Year: {previous}</div>
        <div className="flex items-center gap-2 text-sm">
          <span className={`font-semibold ${positive ? 'text-emerald-400' : 'text-rose-400'}`}>
            {growth == null ? 'N/A' : `${positive ? '▲' : '▼'} ${Math.abs(growth).toFixed(1)}%`}
          </span>
          {extra ? <span className="rounded-full bg-slate-800 px-3 py-1 text-xs text-slate-300">{extra}</span> : null}
        </div>
      </div>
    </div>
  );
};

const Section = ({
  region,
  summary,
}: {
  region: RegionSummary;
  summary: ReturnType<typeof computeRegionMetrics>;
}) => {
  const chartData = createChartData(region);
  return (
    <section className="rounded-[32px] border border-slate-700 bg-slate-900/90 p-6 shadow-xl shadow-slate-950/20">
      <div className="mb-6 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
        <div>
          <h2 className="text-2xl font-semibold text-white">{region.name} VIC</h2>
          <p className="text-slate-400">Performance summary and revenue comparison.</p>
        </div>
      </div>

      <div className="grid gap-4 xl:grid-cols-[1fr_1fr]">
        <div className="grid gap-4">
          <Card title="Standard Inspection">
            <KPIValue
              label="Revenue"
              current={formatAmount(summary.standard.revenueTY)}
              previous={formatAmount(summary.standard.revenueLY)}
              growth={summary.standard.growth}
              extra={region.name.toLowerCase().includes('northern') ? `Share TY ${formatPercent(summary.standard.carParkShareTY)}` : undefined}
            />
            {region.name.toLowerCase().includes('northern') ? (
              <div className="rounded-2xl bg-slate-950/80 p-4 border border-slate-800 text-sm text-slate-300">
                <div>Car Park Share This Year: {formatPercent(summary.standard.carParkShareTY)}</div>
                <div>Car Park Share Last Year: {formatPercent(summary.standard.carParkShareLY)}</div>
              </div>
            ) : null}
          </Card>

          <Card title="Premium Inspection">
            <KPIValue
              label="Revenue"
              current={formatAmount(summary.premium.revenueTY)}
              previous={formatAmount(summary.premium.revenueLY)}
              growth={summary.premium.growth}
              extra={formatPercent(summary.premium.conversion)}
            />
          </Card>

          <Card title="Mobile Inspection (Total)">
            <KPIValue
              label="Revenue"
              current={formatAmount(summary.mobile.revenueTY)}
              previous={formatAmount(summary.mobile.revenueLY)}
              growth={summary.mobile.growth}
              extra={`B2B Conv ${formatPercent(summary.mobile.b2bConversion)}`}
            />
          </Card>

          <Card title="Cross Selling">
            <KPIValue
              label="Revenue"
              current={formatAmount(summary.cross.revenueTY)}
              previous={formatAmount(summary.cross.revenueLY)}
              growth={summary.cross.growth}
              extra={formatPercent(summary.cross.conversion)}
            />
          </Card>

          <Card title="Non-Mandatory Services (Total)">
            <KPIValue
              label="Revenue"
              current={formatAmount(summary.nonMandatory.revenueTY)}
              previous={formatAmount(summary.nonMandatory.revenueLY)}
              growth={summary.nonMandatory.growth}
            />
          </Card>
        </div>

        <div className="rounded-3xl border border-slate-700 bg-slate-950/80 p-5">
          <div className="mb-4 text-sm uppercase tracking-[0.24em] text-slate-400">Revenue comparison</div>
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={chartData} margin={{ top: 10, right: 10, left: -18, bottom: 0 }}>
              <CartesianGrid stroke="#334155" vertical={false} />
              <XAxis dataKey="name" tick={{ fill: '#94a3b8', fontSize: 12 }} tickLine={false} axisLine={false} />
              <YAxis tick={{ fill: '#94a3b8', fontSize: 12 }} tickLine={false} axisLine={false} />
              <Tooltip
                contentStyle={{ background: '#0f172a', border: '1px solid #334155', borderRadius: 16 }}
                formatter={(value: number) => formatAmount(value)}
              />
              <Bar dataKey="lastYear" fill="#334155" radius={[8, 8, 0, 0]} />
              <Bar dataKey="thisYear" fill="#38bdf8" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="mt-6 overflow-x-auto rounded-3xl border border-slate-700 bg-slate-900/90 p-4">
        <div className="mb-4 text-sm uppercase tracking-[0.24em] text-slate-400">Summary table</div>
        <table className="min-w-full border-separate border-spacing-0 text-sm">
          <thead>
            <tr className="bg-slate-950 text-left text-slate-300">
              <th className="border-b border-slate-800 p-3">Service</th>
              <th className="border-b border-slate-800 p-3">TY Revenue</th>
              <th className="border-b border-slate-800 p-3">LY Revenue</th>
              <th className="border-b border-slate-800 p-3">Growth</th>
              <th className="border-b border-slate-800 p-3">Conversion</th>
            </tr>
          </thead>
          <tbody>
            {['Standard Inspection', 'Premium Inspection', 'Mobile Inspection', 'Cross Selling', 'Non-Mandatory Services'].map((label) => {
              const rows = region.services[label] ?? [];
              const { totalTY, totalLY } = getAggValues(rows);
              const growth = computeGrowth(totalTY, totalLY);
              const conversion = label === 'Premium Inspection'
                ? summary.premium.conversion
                : label === 'Mobile Inspection'
                  ? summary.mobile.b2bConversion
                  : label === 'Cross Selling'
                    ? summary.cross.conversion
                    : undefined;
              return (
                <tr key={label} className="border-b border-slate-800 text-slate-200 last:border-0">
                  <td className="p-3 font-medium">{label}</td>
                  <td className="p-3">{formatAmount(totalTY)}</td>
                  <td className="p-3">{formatAmount(totalLY)}</td>
                  <td className="p-3">{growth == null ? 'N/A' : `${growth.toFixed(1)}%`}</td>
                  <td className="p-3">{conversion == null ? '-' : formatPercent(conversion)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
};

const App = () => {
  const [rawText, setRawText] = useState(SAMPLE_PLACEHOLDER);
  const [regions, setRegions] = useState<RegionSummary[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [selectedFileName, setSelectedFileName] = useState<string | null>(null);

  const recommendations = useMemo(() => buildRecommendations(regions), [regions]);

  const handleAnalyze = () => {
    setError(null);
    const rows = parseRows(rawText);
    if (!rows.length) {
      setError('No valid rows detected. Please paste CSV or tabular text with headers.');
      setRegions([]);
      return;
    }
    const summary = buildSummary(rows);
    if (!summary.length) {
      setError('Could not parse the input data. Please check the headers and formatting.');
      setRegions([]);
      return;
    }
    setRegions(summary);
  };

  const handleReset = () => {
    setRawText(SAMPLE_PLACEHOLDER);
    setRegions([]);
    setError(null);
    setSelectedFileName(null);
  };

  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const file = event.target.files?.[0];
    if (!file) return;
    setSelectedFileName(file.name);

    try {
      const text = file.name.match(/\.xls[x]?$/i) || file.type.includes('spreadsheet')
        ? await parseXlsxFile(file)
        : await file.text();
      setRawText(text);
      const rows = parseRows(text);
      if (!rows.length) {
        setError('No valid rows detected in uploaded file. Please check the format.');
        setRegions([]);
        return;
      }
      setRegions(buildSummary(rows));
    } catch (uploadError) {
      setError('Failed to parse uploaded file. Please upload XLS, XLSX, CSV, or TSV.');
      setRegions([]);
    }
  };

  const handleDownload = () => {
    if (!regions.length) return;
    const csv = buildExportCsv(regions);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'vic-kpi-summary.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  };

  return (
    <div className="min-h-screen bg-slate-950 px-4 py-6 sm:px-6 lg:px-10">
      <div className="mx-auto flex max-w-7xl flex-col gap-6">
        <header className="rounded-[36px] border border-slate-700 bg-slate-900/90 p-6 shadow-xl shadow-slate-950/20">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-sm uppercase tracking-[0.32em] text-slate-400">ADNOC Distribution</p>
              <h1 className="mt-2 text-3xl font-semibold text-white sm:text-4xl">VIC KPI Guardian</h1>
              <p className="mt-3 max-w-2xl text-slate-300">Paste raw sales data, click Analyze, and review Abu Dhabi and Northern Emirates VIC performance instantly.</p>
            </div>
            <div className="flex flex-wrap gap-3">
              <button
                onClick={handleAnalyze}
                className="rounded-2xl bg-sky-500 px-5 py-3 text-sm font-semibold text-slate-950 transition hover:bg-sky-400"
              >
                Analyze
              </button>
              <button
                onClick={handleReset}
                className="rounded-2xl border border-slate-700 bg-slate-800 px-5 py-3 text-sm text-slate-200 transition hover:border-slate-500"
              >
                Reset
              </button>
              <button
                onClick={handleDownload}
                disabled={!regions.length}
                className="rounded-2xl bg-slate-700 px-5 py-3 text-sm text-slate-100 transition hover:bg-slate-600 disabled:cursor-not-allowed disabled:opacity-50"
              >
                Download Summary
              </button>
            </div>
          </div>
        </header>

        <section className="rounded-[36px] border border-slate-700 bg-slate-900/90 p-6 shadow-xl shadow-slate-950/20">
          <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-white">Raw Sales Data Input</h2>
              <p className="text-slate-400">Paste CSV, copied rows, or plain text. The parser will detect columns and calculate KPI summaries.</p>
            </div>
          </div>
          <textarea
            className="min-h-[320px] w-full rounded-3xl border border-slate-700 bg-slate-950/95 p-4 text-sm text-slate-100 shadow-inner outline-none focus:border-sky-400 focus:ring-4 focus:ring-sky-500/20"
            value={rawText}
            onChange={(event) => setRawText(event.target.value)}
            placeholder={SAMPLE_PLACEHOLDER}
          />
          <div className="mt-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <label className="group inline-flex cursor-pointer items-center gap-3 rounded-3xl border border-slate-700 bg-slate-950/80 px-4 py-3 text-sm text-slate-200 transition hover:border-slate-500">
              <input
                type="file"
                accept=".xls,.xlsx,.csv,.tsv"
                className="hidden"
                onChange={handleFileUpload}
              />
              <span>Upload XLS/XLSX/CSV</span>
            </label>
            {selectedFileName ? (
              <div className="text-sm text-slate-400">Selected file: {selectedFileName}</div>
            ) : null}
          </div>
          {error ? <p className="mt-4 rounded-2xl bg-rose-500/10 px-4 py-3 text-sm text-rose-200">{error}</p> : null}
        </section>

        {regions.length ? (
          <div className="grid gap-6">
            <div className="grid gap-6 xl:grid-cols-2">
              {['Abu Dhabi', 'Northern Emirates'].map((regionName) => {
                const region = regions.find((item) => item.name.toLowerCase().includes(regionName.toLowerCase()));
                return region ? (
                  <Section key={region.name} region={region} summary={computeRegionMetrics(region)} />
                ) : null;
              })}
            </div>

            <section className="rounded-[32px] border border-slate-700 bg-slate-900/90 p-6 shadow-xl shadow-slate-950/20">
              <div className="mb-4 text-xl font-semibold text-white">Patterns & Recommendations</div>
              <div className="grid gap-3">
                {recommendations.length ? (
                  recommendations.map((item, index) => (
                    <div
                      key={`${item.type}-${index}`}
                      className={`rounded-3xl border p-4 text-sm ${
                        item.type === 'warning'
                          ? 'border-rose-600 bg-rose-500/10 text-rose-100'
                          : 'border-sky-600 bg-sky-500/10 text-sky-100'
                      }`}
                    >
                      {item.message}
                    </div>
                  ))
                ) : (
                  <div className="rounded-3xl border border-slate-700 bg-slate-950/80 p-4 text-slate-300">No immediate flags detected. Review the results for further optimization opportunities.</div>
                )}
              </div>
            </section>
          </div>
        ) : (
          <section className="rounded-[32px] border border-dashed border-slate-700 bg-slate-900/80 p-6 text-slate-400">
            <p>Use the Analyze button to generate the VIC KPI dashboard after pasting sales data.</p>
          </section>
        )}
      </div>
    </div>
  );
};

export default App;
