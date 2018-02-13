script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.CustomDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.dialogs.ProcessingStatusDialog"
java_import "com.nuix.nx.digest.DigestHelper"
java_import "com.nuix.nx.controls.models.Choice"

java_import java.lang.Runtime
java_import java.io.BufferedReader
java_import java.io.InputStreamReader

require 'csv'
require 'fileutils'
require 'json'

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

# Search case for registry hive items
search_query = "mime-type:application/vnd.ms-registry AND NOT name:( sav OR old OR log* ) AND " +
"((path-name:Windows/System32/config AND name:( sam OR security OR software OR system OR default )) OR " +
"(path-name:( AppData/Roaming/Microsoft/Windows OR AppData/Local/Microsoft/Windows OR \"Local Settings/Application Data/Microsoft/Windows\") AND name:USRCLASS.DAT) OR " +
"(path-name:( (Users OR \"Documents and Settings\" OR Windows/system32/config/systemprofile OR Windows/ServiceProfiles/LocalService OR Windows/ServiceProfiles/NetworkService) AND NOT windows/users ) AND name:ntuser.dat) OR " +
"(path-name:Windows/AppCompat/Programs AND name:Amcache.hve))"

# Load HiveProfileMap
hive_profile = JSON.parse(File.read(File.join(script_directory,"HiveProfileMap.json")))
	
# Setup the dialog
dialog = TabbedCustomDialog.new

reg_ripper_tab = dialog.addTab("reg_ripper_tab", "RegRipper")
reg_ripper_tab.appendDirectoryChooser("rr_path", "RegRipper Installation Path")
reg_ripper_tab.appendDirectoryChooser("export_path", "Export Path")
reg_ripper_tab.appendCheckBox("delete_export", "Delete export on completion?", false)
reg_ripper_tab.appendDirectoryChooser("output_path", "Output Path")

# Validation
dialog.validateBeforeClosing do |values|
	if values["rr_path"].strip.empty?
		CommonDialogs.showWarning("Please provide the RegRipper Installation Path.")
		next false
	else
		rip = java.io.File.new(File.join(values["rr_path"]),"rip.exe")
		if !rip.exists || !rip.isFile
			CommonDialogs.showWarning("RegRipper Installation Path is invalid.")
			next false
		end
  end
	
	if values["export_path"].strip.empty?
		CommonDialogs.showWarning("Please provide the Export Path.")
		next false
	else
		export_dir = java.io.File.new(values["export_path"].strip)
		if !export_dir.exists || !export_dir.isDirectory
			CommonDialogs.showWarning("Export Path is invalid.")
			next false
		end
	end
	
	if values["output_path"].strip.empty?
		CommonDialogs.showWarning("Please provide the Output Path.")
		next false
	else
		output_dir = java.io.File.new(values["output_path"].strip)
		if !output_dir.exists || !output_dir.isDirectory
			CommonDialogs.showWarning("Output Path is invalid.")
			next false
		end
	end
	
	next true
end

dialog.display

# Convenience method for running a command string in the OS
#
# @param command [String] Command string to execute
# @param use_shell [Boolean] When true, will pipe command through CMD /S /C to enable shell features
# @param working_dir [String] The working direcotry of the subprocess
def run(command,use_shell=true,working_dir)
	# Necessary if command take advantage of any shell features such as
	# IO pipining
	if use_shell
		command = "cmd /S /C \"#{command}\""
	end

	begin
		puts "Executing: #{command}"
		p = Runtime.getRuntime.exec(command,[].to_java(:string),java.io.File.new(working_dir))
		
		# RegRipper sends output to the error stream. This must be read in order to prevent the process from locking
		std_err_reader = BufferedReader.new(InputStreamReader.new(p.getErrorStream))
		while ((line = std_err_reader.readLine()).nil? == false)
			#puts line
		end
		
		p.waitFor
		puts "Execution completed:"
		reader = BufferedReader.new(InputStreamReader.new(p.getInputStream))
		while ((line = reader.readLine()).nil? == false)
			puts line
		end
	rescue Exception => e
		puts e.message
		puts e.backtrace.inspect
	ensure
		p.destroy
	end
end
	
if dialog.getDialogResult == true
	input = dialog.toMap
	
	rr_install_path = input["rr_path"]
	rr_path = File.join(rr_install_path,"rip.exe")
	export_path = input["export_path"]
	output_path = input["output_path"]
	delete_export = input["delete_export"]
	
	# Keep track of report files generated to add index to file name in case of duplicate report names
	report_files = Hash.new
	
	summary_report = Array.new
	
	registry_items = $currentCase.search(search_query)
	
	if registry_items.size < 1
		CommonDialogs.showWarning("Case does not contain Registry items.")
	else
		# Create MDP for loadfile
		mdp = $utilities.getMetadataProfileStore.createMetadataProfile.addMetadata("itemGuid") do |item|
			next item.getGuid
		end
		
		ProgressDialog.forBlock do |pd|
			pd.setTitle("RegRipper")
			
			# Ensure propper logging
			pd.onMessageLogged do |message|
				puts message
			end
			
			# 8 export stages + regripper stage
			# Export stages: native, store_email_fixup, numbering, file_naming, set_file_times, digest, create_production_set, load_files
			main_progress = 0
			pd.setMainProgress(main_progress, registry_items.size * 9)
			pd.setMainProgressVisible(true)
			pd.setMainStatusAndLogIt("Exporting items: Step 1/2")
			
			sub_progress = 0
			pd.setSubProgress(sub_progress,registry_items.size * 8)
			pd.setSubProgressVisible(true)
			pd.setSubStatusAndLogIt("Exporting items...")
			
			# Export registry files
			export_dir = File.join(export_path,"RegRipper_Export_" + Time.now.strftime("%Y%m%d_%H%M%S"))
			FileUtils.mkdir(File.join(export_dir))
			exporter = $utilities.createBatchExporter(export_dir)
			exporter.addProduct("native",{"naming" => "item_name_with_path"})
			exporter.addLoadFile("csv_summary_report", {"metadataProfile" => mdp})
			exporter.whenItemEventOccurs do |callback|
				main_progress += 1
				sub_progress += 1
				pd.setMainProgress(main_progress)
				pd.setSubProgress(sub_progress)
			end
			exporter.exportItems(registry_items)
			pd.logMessage("Finished Exporting.")
			
			pd.setMainStatusAndLogIt("Generating Reports: Step 2/2")
			sub_progress = 0
			pd.setSubProgress(sub_progress, registry_items.size)
			pd.setSubStatusAndLogIt("Generating reports...")
			
			CSV.foreach(File.join(export_dir,"Report.csv"), {:headers => true, :skip_blanks => true}) do |row|
				break if pd.abortWasRequested
				
				main_progress += 1
				sub_progress += 1
				pd.setMainProgress(main_progress)
				pd.setSubProgress(sub_progress)
				pd.setSubStatus("Generating report #{sub_progress}/#{registry_items.size}")
				
				evidence_container = row[1].split("/")[0]
				hive_file_path = File.join(export_dir,row[0])
				exported_hive_file = File.join(export_dir,row[1])
				hive_name = File.basename(hive_file_path).downcase
				report_name = evidence_container
				
				if hive_file_path.downcase.include? "regback"
					report_name << "_RegBack"
				end
				
				# Attempt to append username to report filename
				if hive_name == "ntuser.dat" || hive_name == "usrclass.dat"
					hive_dir_array = File.dirname(hive_file_path).split("/")
					if hive_dir_array.length > 0
						user_folder = nil
						["Users", "Documents and Settings", "ServiceProfiles"].each do |parent_folder|
							parent_index = hive_dir_array.index(parent_folder)
							if parent_index && hive_dir_array.size >= parent_index + 1 && hive_dir_array[parent_index + 1]
								user_folder = "#{hive_dir_array[parent_index + 1]}"
							end
						end
						if user_folder
							report_name << "_#{user_folder}"
						else
							report_name << "_#{hive_dir_array[hive_dir_array.length - 1]}"
						end
					end
				end
				
				report_name << "_#{hive_name}"
				
				# Check for existing report, append index to report name if found
				if report_files.has_key?(report_name)
					report_files[report_name] = report_files[report_name] + 1
					report_name << "_#{report_files[report_name]}"
				else
					report_files[report_name] = 0
				end
				
				output_file = File.join(output_path, "#{report_name}.txt")
				
				# Determine which RegRipper profile to use based on HiveProfileMap.json settings
				rr_profile = "all"
				if hive_profile.has_key?(hive_name)
					rr_profile = hive_profile[hive_name]
				end
				
				pd.logMessage("Generating report #{sub_progress}/#{registry_items.size}: #{output_file}")
				#command = "\"C:/Program Files/RegRipper2.8-master/rip.exe\" -r \"c:/evidence/registry files/NTUSER.DAT\" -f ntuser >\"c:/evidence/registry files/ntuser.txt\""
				command = "\"#{rr_path}\" -r \"#{exported_hive_file}\" -f #{rr_profile} >\"#{output_file}\""

				run(command,true,rr_install_path)
				
				summary_report << [output_file,exported_hive_file,row[2]]
			end
			
			# Write summary report
			CSV.open(File.join(output_path,"summary_report.csv"), "wb", {:force_quotes => true}) do |csv|
				csv << ["Report file", "Hive File", "Item GUID"]
				summary_report.each do |report|
					csv << report
				end
			end
			
			if delete_export
				pd.logMessage("Deleting export directory.")
				FileUtils.remove_dir(export_dir,true)
			end
			
			if pd.abortWasRequested
				pd.logMessage("Aborting...")
			end
			
			pd.logMessage("Completed!")
		end
	end
end
