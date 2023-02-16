#include <Origin.h>

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



static void add_to_list (string name, vector <string> &names, vector<int> &counts)
{
	int idx = names.Find(name);
	if (idx == -1) {
		names.Add(name);
		counts.Add(1);
	} else {
		counts[idx]++;
	}
}



static bool has_letter (string long_name)
{
	int length = strlen(long_name);
	if (length < 2)
		return false;
	char penultimate = long_name[length - 2], last = long_name[length - 1];

	return (penultimate == '_') && (last >= 'a') && (last <= 'z');
}



static void add_letters (vector <string> names, vector<int> counts)
{
// we do not want to add a '_a' at the end of unique names
	for (int i = 0; i < names.GetSize(); i++)
		if (counts[i] == 1)
			counts[i] = 0;

	foreach (const PageBase pagebase in Project.ActiveFolder().Pages) {
		string long_name = pagebase.GetLongName();
		if (has_letter(long_name))
			continue;

		int idx = names.Find(long_name);
		if (idx == -1) {
			printf("error: no files named %s in the names vector\n", long_name);
			continue;
		}
		int count = counts[idx];
		if (count <= 0)
			continue;

		char letter = 'a' + count - 1;
		long_name += "_" + letter;
		pagebase.SetLongName(long_name, false, true);
		counts[idx]--;
	}
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

	if (streq(folder_name, Project.RootFolder.GetName())) {
		printf("The active folder is the root folder, nothing to do.");
		return;
	}

	vector <string> names;
	vector <int> counts;

	foreach (const PageBase pagebase in folder.Pages) {
		if (pagebase.GetType() != EXIST_WKS)
			continue;

		const string cur_long_name = pagebase.GetLongName();
		printf("\n\nworksheet:\tshort name = '%s' ; long name = '%s'\n",
			pagebase.GetName(), cur_long_name
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
			add_to_list(new_long_name, names, counts);

			if (has_letter(cur_long_name))
				new_long_name += "_" + (char)str_end_char(cur_long_name);

			printf("new name:\t\"%s\"\n", new_long_name);
			if (!pagebase.SetLongName(new_long_name, false, true))
				printf("Error, The worksheet was not renamed.");
		} catch (int errcode) {
			print_missing_param(errcode);
			printf("The worksheet was not renamed.");
		}
	}
	add_letters(names, counts);
}






// for debugging

static void print_vector (vector <string> strings)
{
	int size = strings.GetSize();
	if (size == 0) {
		printf("{}\n");
		return;
	}
	printf("{");
	for (int i = 0; i < size - 1; i++)
		printf("\"%.30s\", ", strings[i]);
	printf("\"%.30s\"}\n", strings[size - 1]);
}


static void print_vector (vector <int> numbers)
{
	int size = numbers.GetSize();
	if (size == 0) {
		printf("{}\n");
		return;
	}
	printf("{");
	for (int i = 0; i < size - 1; i++)
		printf("\"%d\", ", numbers[i]);
	printf("\"%d\"}\n", numbers[size - 1]);
}