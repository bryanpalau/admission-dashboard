import React, { useState, useEffect } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, ScatterChart, Scatter, ZAxis,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar
} from 'recharts';
import * as XLSX from 'xlsx';

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d', '#ffc658', '#8dd1e1', '#a4de6c', '#d0ed57'];

const Dashboard = () => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('overview');
  const [selectedSchoolSize, setSelectedSchoolSize] = useState('all');

  const cleanData = (rawData) => {
    return rawData.map(row => {
      let conversionRate;
      if (typeof row["Conversion Rate (%)"] === 'string') {
        conversionRate = parseFloat(row["Conversion Rate (%)"].replace('%', ''));
      } else {
        conversionRate = row["Conversion Rate (%)"];
        if (conversionRate > 1) {
          conversionRate = conversionRate;
        } else {
          conversionRate = conversionRate * 100;
        }
      }
      
      return {
        school: row.School,
        applications: typeof row.Appllication === 'number' ? row.Appllication : null,
        admissions: row.Admission,
        conversionRate: conversionRate
      };
    }).filter(item => 
      item.applications !== null && 
      !isNaN(item.applications) && 
      !isNaN(item.admissions) && 
      !isNaN(item.conversionRate)
    );
  };

  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        
        const response = await fetch('/admission-dashboard/data/Admission_Conversion_Rates_20250411.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(response, {
          cellStyles: true,
          cellFormulas: true,
          cellDates: true,
          cellNF: true,
          sheetStubs: true
        });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet);

        // Clean and transform the data
        setData(cleanData(rawData));
        setLoading(false);
      } catch (error) {
        console.error("Error loading data:", error);
        setLoading(false);
      }
    };

    loadData();
  }, []);

  const getFilteredData = () => {
    if (selectedSchoolSize === 'all') return data;
    
    if (selectedSchoolSize === 'small') {
      return data.filter(item => item.applications < 20);
    } else if (selectedSchoolSize === 'medium') {
      return data.filter(item => item.applications >= 20 && item.applications < 50);
    } else if (selectedSchoolSize === 'large') {
      return data.filter(item => item.applications >= 50 && item.applications < 100);
    } else if (selectedSchoolSize === 'xlarge') {
      return data.filter(item => item.applications >= 100);
    }
    
    return data;
  };

  const topSchoolsByApplications = () => {
    return [...data]
      .sort((a, b) => b.applications - a.applications)
      .slice(0, 10)
      .map(item => ({
        ...item,
        name: item.school.length > 20 ? item.school.substring(0, 20) + '...' : item.school
      }));
  };

  const topSchoolsByConversionRate = () => {
    return [...data]
      .filter(item => item.applications >= 20) // Only schools with significant applications
      .sort((a, b) => b.conversionRate - a.conversionRate)
      .slice(0, 10)
      .map(item => ({
        ...item,
        name: item.school.length > 20 ? item.school.substring(0, 20) + '...' : item.school
      }));
  };

  const getConversionRateDistribution = () => {
    const bins = [
      { name: '0-20%', value: 0 },
      { name: '20-40%', value: 0 },
      { name: '40-60%', value: 0 },
      { name: '60-80%', value: 0 },
      { name: '80-100%', value: 0 }
    ];

    getFilteredData().forEach(school => {
      const rate = school.conversionRate;
      if (rate < 20) bins[0].value++;
      else if (rate < 40) bins[1].value++;
      else if (rate < 60) bins[2].value++;
      else if (rate < 80) bins[3].value++;
      else bins[4].value++;
    });

    return bins;
  };

  const getSchoolSizeDistribution = () => {
    return [
      { 
        name: '1-20 Applications', 
        value: data.filter(s => s.applications >= 1 && s.applications < 20).length 
      },
      { 
        name: '20-50 Applications', 
        value: data.filter(s => s.applications >= 20 && s.applications < 50).length 
      },
      { 
        name: '50-100 Applications', 
        value: data.filter(s => s.applications >= 50 && s.applications < 100).length 
      },
      { 
        name: '100+ Applications', 
        value: data.filter(s => s.applications >= 100).length 
      }
    ];
  };

  const getScatterData = () => {
    return getFilteredData().map(item => ({
      x: item.applications,
      y: item.conversionRate,
      z: item.admissions,
      name: item.school
    }));
  };

  const getTotalApplicationsAndAdmissions = () => {
    const filteredData = getFilteredData();
    const totalApplications = filteredData.reduce((sum, item) => sum + item.applications, 0);
    const totalAdmissions = filteredData.reduce((sum, item) => sum + item.admissions, 0);
    const overallRate = (totalAdmissions / totalApplications * 100).toFixed(2);
    
    return {
      applications: totalApplications,
      admissions: totalAdmissions,
      rate: overallRate
    };
  };

  const CustomTooltip = ({ active, payload, label }) => {
    if (active && payload && payload.length) {
      return (
        <div className="bg-white p-4 border border-gray-200 rounded shadow-lg">
          <p className="font-bold">{payload[0].payload.name}</p>
          <p>Applications: {payload[0].payload.applications}</p>
          <p>Admissions: {payload[0].payload.admissions}</p>
          <p>Conversion Rate: {payload[0].payload.conversionRate.toFixed(2)}%</p>
        </div>
      );
    }
    return null;
  };

  const ScatterTooltip = ({ active, payload }) => {
    if (active && payload && payload.length) {
      return (
        <div className="bg-white p-4 border border-gray-200 rounded shadow-lg">
          <p className="font-bold">{payload[0].payload.name}</p>
          <p>Applications: {payload[0].payload.x}</p>
          <p>Conversion Rate: {payload[0].payload.y.toFixed(2)}%</p>
          <p>Admissions: {payload[0].payload.z}</p>
        </div>
      );
    }
    return null;
  };

  const stats = getTotalApplicationsAndAdmissions();

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="text-2xl text-gray-600">Loading data...</div>
      </div>
    );
  }

  return (
    <div className="bg-gray-50 p-6 rounded-lg">
      <h1 className="text-3xl font-bold text-gray-800 mb-6">School Admission Conversion Dashboard</h1>
      
      {/* Filter Controls */}
      <div className="mb-6 bg-white p-4 rounded-lg shadow">
        <div className="flex flex-wrap gap-4 items-center">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Filter by School Size:</label>
            <select 
              className="bg-white border border-gray-300 rounded px-3 py-2"
              value={selectedSchoolSize}
              onChange={e => setSelectedSchoolSize(e.target.value)}
            >
              <option value="all">All Schools</option>
              <option value="small">Small (1-19 Applications)</option>
              <option value="medium">Medium (20-49 Applications)</option>
              <option value="large">Large (50-99 Applications)</option>
              <option value="xlarge">X-Large (100+ Applications)</option>
            </select>
          </div>
          
          <div className="flex-grow"></div>
          
          {/* Tabs */}
          <div className="flex border rounded overflow-hidden">
            <button 
              className={`px-4 py-2 ${activeTab === 'overview' ? 'bg-blue-500 text-white' : 'bg-white'}`}
              onClick={() => setActiveTab('overview')}
            >
              Overview
            </button>
            <button 
              className={`px-4 py-2 ${activeTab === 'schools' ? 'bg-blue-500 text-white' : 'bg-white'}`}
              onClick={() => setActiveTab('schools')}
            >
              Top Schools
            </button>
            <button 
              className={`px-4 py-2 ${activeTab === 'analysis' ? 'bg-blue-500 text-white' : 'bg-white'}`}
              onClick={() => setActiveTab('analysis')}
            >
              Analysis
            </button>
          </div>
        </div>
      </div>
      
      {/* Stats Cards */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-semibold text-gray-700 mb-2">Total Applications</h3>
          <p className="text-3xl font-bold text-blue-600">{stats.applications.toLocaleString()}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-semibold text-gray-700 mb-2">Total Admissions</h3>
          <p className="text-3xl font-bold text-green-600">{stats.admissions.toLocaleString()}</p>
        </div>
        <div className="bg-white p-6 rounded-lg shadow">
          <h3 className="text-lg font-semibold text-gray-700 mb-2">Overall Conversion Rate</h3>
          <p className="text-3xl font-bold text-purple-600">{stats.rate}%</p>
        </div>
      </div>
      
      {/* Main Content */}
      {activeTab === 'overview' && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">School Size Distribution</h2>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={getSchoolSizeDistribution()}
                    cx="50%"
                    cy="50%"
                    outerRadius={80}
                    fill="#8884d8"
                    dataKey="value"
                    nameKey="name"
                    label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  >
                    {getSchoolSizeDistribution().map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
          
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Conversion Rate Distribution</h2>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={getConversionRateDistribution()}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" />
                  <YAxis />
                  <Tooltip />
                  <Legend />
                  <Bar dataKey="value" name="Number of Schools" fill="#8884d8" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          
          <div className="bg-white p-4 rounded-lg shadow lg:col-span-2">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Applications vs. Conversion Rate</h2>
            <div className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 20 }}>
                  <CartesianGrid />
                  <XAxis type="number" dataKey="x" name="Applications" />
                  <YAxis type="number" dataKey="y" name="Conversion Rate (%)" domain={[0, 100]} />
                  <ZAxis type="number" dataKey="z" range={[50, 500]} name="Admissions" />
                  <Tooltip content={<ScatterTooltip />} />
                  <Scatter name="Schools" data={getScatterData()} fill="#8884d8" />
                </ScatterChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}
      
      {activeTab === 'schools' && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Top Schools by Applications</h2>
            <div className="h-96">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={topSchoolsByApplications()}
                  layout="vertical"
                  margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis type="number" />
                  <YAxis type="category" dataKey="name" width={150} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend />
                  <Bar dataKey="applications" name="Applications" fill="#0088FE" />
                  <Bar dataKey="admissions" name="Admissions" fill="#00C49F" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Top Schools by Conversion Rate</h2>
            <div className="h-96">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={topSchoolsByConversionRate()}
                  layout="vertical"
                  margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis type="number" domain={[0, 100]} />
                  <YAxis type="category" dataKey="name" width={150} />
                  <Tooltip content={<CustomTooltip />} />
                  <Legend />
                  <Bar dataKey="conversionRate" name="Conversion Rate (%)" fill="#FFBB28" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          
          <div className="bg-white p-4 rounded-lg shadow lg:col-span-2">
            <h2 className="text-xl font-bold text-gray-800 mb-4">All Schools Comparison</h2>
            <div className="h-96 overflow-auto">
              <table className="min-w-full bg-white">
                <thead className="bg-gray-100">
                  <tr>
                    <th className="py-2 px-4 border-b text-left">School</th>
                    <th className="py-2 px-4 border-b text-right">Applications</th>
                    <th className="py-2 px-4 border-b text-right">Admissions</th>
                    <th className="py-2 px-4 border-b text-right">Conversion Rate</th>
                  </tr>
                </thead>
                <tbody>
                  {getFilteredData().sort((a, b) => b.applications - a.applications).map((school, index) => (
                    <tr key={index} className={index % 2 === 0 ? 'bg-gray-50' : 'bg-white'}>
                      <td className="py-2 px-4 border-b">{school.school}</td>
                      <td className="py-2 px-4 border-b text-right">{school.applications}</td>
                      <td className="py-2 px-4 border-b text-right">{school.admissions}</td>
                      <td className="py-2 px-4 border-b text-right">{school.conversionRate.toFixed(2)}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
      
      {activeTab === 'analysis' && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Applications vs. Admissions Trend</h2>
            <div className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart
                  data={topSchoolsByApplications()}
                  margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" />
                  <YAxis />
                  <Tooltip />
                  <Legend />
                  <Line type="monotone" dataKey="applications" name="Applications" stroke="#0088FE" activeDot={{ r: 8 }} />
                  <Line type="monotone" dataKey="admissions" name="Admissions" stroke="#00C49F" />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="bg-white p-4 rounded-lg shadow">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Radar Analysis</h2>
            <div className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <RadarChart outerRadius={90} data={topSchoolsByApplications().slice(0, 5)}>
                  <PolarGrid />
                  <PolarAngleAxis dataKey="name" />
                  <PolarRadiusAxis angle={30} domain={[0, 150]} />
                  <Radar name="Applications" dataKey="applications" stroke="#8884d8" fill="#8884d8" fillOpacity={0.6} />
                  <Radar name="Admissions" dataKey="admissions" stroke="#82ca9d" fill="#82ca9d" fillOpacity={0.6} />
                  <Legend />
                  <Tooltip />
                </RadarChart>
              </ResponsiveContainer>
            </div>
          </div>
          
          <div className="bg-white p-4 rounded-lg shadow lg:col-span-2">
            <h2 className="text-xl font-bold text-gray-800 mb-4">Performance Summary</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <h3 className="font-semibold text-gray-700 mb-2">Key Insights:</h3>
                <ul className="list-disc pl-5 space-y-2">
                  <li>Top performing school by conversion rate: {
                    [...data].sort((a, b) => b.conversionRate - a.conversionRate)[0]?.school || 'N/A'
                  } ({[...data].sort((a, b) => b.conversionRate - a.conversionRate)[0]?.conversionRate.toFixed(2) || 'N/A'}%)</li>
                  <li>School with most applications: {
                    [...data].sort((a, b) => b.applications - a.applications)[0]?.school || 'N/A'
                  } ({[...data].sort((a, b) => b.applications - a.applications)[0]?.applications || 'N/A'})</li>
                  <li>Average conversion rate: {
                    (data.reduce((sum, item) => sum + item.conversionRate, 0) / data.length).toFixed(2)
                  }%</li>
                  <li>Most schools ({
                    Math.round(data.filter(s => s.conversionRate >= 60 && s.conversionRate < 80).length / data.length * 100)
                  }%) have conversion rates between 60-80%</li>
                </ul>
              </div>
              <div>
                <h3 className="font-semibold text-gray-700 mb-2">Recommendations:</h3>
                <ul className="list-disc pl-5 space-y-2">
                  <li>Focus on schools with high application volumes but below-average conversion rates</li>
                  <li>Study best practices from top converting schools</li>
                  <li>Consider targeted campaigns for schools with low application numbers</li>
                  <li>Investigate any schools with conversion rates below 60% to identify issues</li>
                </ul>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// Add footer with export options
const Footer = () => {
  const handleExportCSV = () => {
    alert("CSV export feature would be implemented here");
  };

  const handleExportPDF = () => {
    alert("PDF export feature would be implemented here");
  };

  const handleSaveImage = () => {
    alert("Image export feature would be implemented here");
  };

  return (
    <div className="mt-8 pt-4 border-t border-gray-200">
      <div className="flex flex-wrap justify-between items-center">
        <div>
          <h3 className="text-lg font-semibold text-gray-700">Export Options</h3>
          <p className="text-sm text-gray-500">Share or download this dashboard in various formats</p>
        </div>
        <div className="flex space-x-2 mt-2">
          <button 
            onClick={handleExportCSV}
            className="bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded"
          >
            Export CSV
          </button>
          <button 
            onClick={handleExportPDF}
            className="bg-green-500 hover:bg-green-600 text-white px-4 py-2 rounded"
          >
            Export PDF
          </button>
          <button 
            onClick={handleSaveImage}
            className="bg-purple-500 hover:bg-purple-600 text-white px-4 py-2 rounded"
          >
            Save as Image
          </button>
        </div>
      </div>
    </div>
  );
};

// Main App component
const App = () => {
  return (
    <div className="max-w-6xl mx-auto px-4 py-8">
      <Dashboard />
      <Footer />
    </div>
  );
};

export default App;