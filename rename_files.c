#include <Origin.h>
#include <Array.h>

typedef enum EXP_TYPE { EXCITATION, EMISSION } EXP_TYPE;



typedef enum MISSING_PARAM {
	NO_PARK,
	NO_EX_SLIT,
	NO_EM_SLIT,
	NO_EXP_TYPE,
	NO_INT_TIME,
} MISSING_PARAM;

static void print_missing_param (MISSING_PARAM param)
{
	string name;
	switch (param) {
	case NO_PARK    : name = "park";             break;
	case NO_EX_SLIT : name = "excitation slit";  break;
	case NO_EM_SLIT : name = "emission slit";    break;
	case NO_EXP_TYPE: name = "experiment type";  break;
	case NO_INT_TIME: name = "integration time"; break;
	default         : name = "(unknown)";        break;
	}
	printf("unable to find the following experimental parameter: %s\n", name);
}



static bool streq (string a, string b)
{
	return !strcmp(a, b);
}



static bool has_decimals (float f)
{
	float decimal_part = rmod(f, 1.0);
	return fabs(decimal_part) < 1e-9;
}



static string build_long_name (string folder_name, EXP_TYPE exp_type, float park, float ex_slit, float em_slit, float integration_time)
{
	char namebuf[256];
	sprintf(namebuf, "%s%s_%.0f_%.*f_%.*f_%.*f",
		(exp_type == EMISSION) ? "" : "Ex_",
		folder_name,
		park,
// for the following fields, if there are no decimals, we do not want the .0
// else, we want 1 decimal
#define precision_and_num(f) has_decimals(f) ? 0 : 1, (f)

		precision_and_num(ex_slit),
		precision_and_num(em_slit),
		precision_and_num(integration_time)
	);
	string name(namebuf);
	return name;
}



static string extract_parameters (string folder_name, string exp_string)
{
	vector<string> lines;
	int nb_lines = str_separate(exp_string, "\r\n", lines);

	char exp_type_str[64] = "";
	// doubles do not work with %f format ...
	float integration_time = -1, park = -1, em_slit = -1, ex_slit = -1;
	EXP_TYPE mode;

	for (int i = 0; i < nb_lines; i++) {
		string str = lines[i];

		if (is_str_match_begin("Experiment Type:", str))
			sscanf(str, "Experiment Type: Spectral Acquisition[%63[^]]]", exp_type_str);
		else if (is_str_match_begin("Integration time:", str))
			sscanf(str, "Integration Time: %fs", &integration_time);
		else if (is_str_match_begin("Park:", str))
			sscanf(str, "Park: %fnm", &park);
		else if (is_str_match_begin("EX1: Excitation", str))
			mode = EXCITATION;
		else if (is_str_match_begin("EM1: Emission", str))
			mode = EMISSION;
		else if (is_str_match_begin("Side Entrance Slit:", str))
			sscanf(str, "Side Entrance Slit: %f nmBandpass", (mode == EMISSION) ? &em_slit : &ex_slit);
	}

	printf("exp_type = %s, park = %.0f, ex_slit = %.1f, em_slit = %.1f, integration_time = %.1f,\n",
		exp_type_str, park, ex_slit, em_slit, integration_time
	);

	// if parameters still have their default value, it means they were not found
	if (streq(exp_type_str, ""))
		throw NO_EXP_TYPE;
	if (integration_time == -1.0)
		throw NO_INT_TIME;
	if (ex_slit == -1.0)
		throw NO_EX_SLIT;
	if (em_slit == -1.0)
		throw NO_EM_SLIT;
	if (park == -1.0)
		throw NO_PARK;

	return build_long_name(
		folder_name,
		streq(exp_type_str, "Emission") ? EMISSION : EXCITATION,
		park, ex_slit, em_slit,
		integration_time
	);
}



string get_creation_date (string short_name)
{
// Page FindPage( LPCSTR lpcszName, int nType = EXIST_WKS, int nAlsoCanbeType = EXIST_MATRIX, BOOL bAlsoCanbeLongName = true )
// nType : The type of page to look for, use EXIST_WKS, EXIST_MATRIX, etc, or 0 if any type
// nAlsoCanbeType : The Or condition when searching, must be -1 if nType is 0, and can never be 0
	Page page = Project.FindPage(short_name, 0, -1, false);
	if (!page.IsValid())
		return "PAGE_ERROR";
	PropertyInfo pageInfo;
/* typedef struct tagPropertyInfo {
	char    szSize[MAXLINE];
	char    szType[MAXLINE];
	char    szState[MAXLINE];
	char    szCreate[MAXLINE];
	char    szModify[MAXLINE];
	char    szLocation[MAXLINE];
	char    szContains[MAXLINE];
}PropertyInfo, *pPropertyInfo; */
	if (page.GetPageInfo(pageInfo))
		return pageInfo.szCreate;
	else
		return "INFO_ERROR";
}



static time_t datestring_to_epoch_time (string datestring)
{
	int day, month, year, hours, minutes;
// format: "02/02/2023 15:11"
	sscanf(datestring, "%2d/%2d/%4d %2d:%2d",
		&day, &month, &year, &hours, &minutes
	);

	struct tm date = {0};
	date.tm_year = year - 1900;
	date.tm_mon = month - 1;
	date.tm_mday = day;
	date.tm_hour = hours;
	date.tm_min = minutes;

	return mktime(&date);
}



typedef struct PageStruct {
	Page page;
	string name;
	time_t creation_time;
} PageStruct;



static void sort_array (Array<PageStruct&> &array)
{
	const int size = array.GetSize();
	vector <int> creation_times; // .Sort() does not work with 64-bit integers like time_t...

	for (int i = 0; i < size; i++) {
		PageStruct& ptr = array.GetAt(i);
		creation_times.Add(ptr.creation_time);
	}

	vector <uint> indexes;
	if (!creation_times.Sort(SORT_DESCENDING, true, indexes, SORTCNTRL_STABLE_ALGORITHM))
		printf("Error: failed to sort array");

	Array<PageStruct&> sorted;
	for (i = 0; i < size; i++)
		sorted.Add(array.GetAt(indexes[i]));
	for (i = 0; i < size; i++)
		array.SetAt(i, sorted.GetAt(i));
}



void rename_files (void)
{
	Folder folder = Project.ActiveFolder();

	printf("\n\n"
		"=================================================""\n"
		"active folder:\t%s"                               "\n"
		"=================================================""\n",
		folder.GetPath()
	);

	const string folder_name = folder.GetName();

	vector <string> names;
	Array<PageStruct&> pagesArray;

	foreach (const PageBase pagebase in folder.Pages) {
		string name = pagebase.GetName(), long_name = pagebase.GetLongName();
		if ((pagebase.GetType() != EXIST_WKS) || is_str_match_begin("NORM", long_name) || is_str_match_begin("STACK", long_name))
			continue;

		string creation_date = get_creation_date(name);

		printf("\n\nworksheet: created %s\tshort name = '%s' ; long name = '%s'\n",
			creation_date, name, long_name
		);

		Page page;
		page = (Page)pagebase;

		Worksheet worksheet = page.Layers("Note");
		if (!worksheet) {
			printf("this worksheet does not have a sheet named 'Note'\n");
			continue;
		}

		vector<string> columns;
		if (!worksheet.Columns(0).GetStringArray(columns)) {
			printf("unable to read the Note\n");
			continue;
		}
		string content = columns[0];

		try {
			string new_long_name = extract_parameters(folder_name, content);
			printf("new name:\t\"%s\"\n", new_long_name);
			names.Add(new_long_name);

			PageStruct *page_struct = new PageStruct;
			page_struct->page = page;
			page_struct->name = new_long_name;
			page_struct->creation_time = datestring_to_epoch_time(creation_date);

			pagesArray.Add(*page_struct);

		} catch (int errcode) {
			print_missing_param(errcode);
			printf("The worksheet was not renamed.");
		}
	}
	sort_array(pagesArray);

	vector<string> unique_names;
	vector<uint> counts;
	count_list(names, unique_names, counts);

	for (int i = 0; i < pagesArray.GetSize(); i++) {
		PageStruct& page_struct = pagesArray.GetAt(i);
		string name = page_struct.name;
		int idx = unique_names.Find(name);
		if (idx == -1)
			printf("could not find name %s in the names vector\n", page_struct.name);
		Page page = page_struct.page;
		uint count = counts[idx];
		if (count != 1)
			name += "-" + (count - 1);

		PropertyInfo info;
		page.GetPageInfo(info);
		printf("renaming: created %s, old name = %s, new name = %s\n", info.szCreate, page.GetLongName(), name);
		if (!page.SetLongName(name, false, true))
			printf("unable to rename page %s (%s)", page.GetName(), page.GetLongName());
		counts[idx]--;
	}
}





// for debugging

static void print_vector (vector <string> strings)
{
	int size = strings.GetSize();
	printf("{");
	for (int i = 0; i < size; i++)
		printf("\"%.30s\"%s", strings[i], (i == size - 1) ? "" : ", ");
}


static void print_vector (vector <int> numbers)
{
	int size = numbers.GetSize();
	printf("{");
	for (int i = 0; i < size; i++)
		printf("\"%d\"%s", numbers[i], (i == size - 1) ? "" : ", ");
}