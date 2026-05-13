###############################################################################
#
#                       S H O P D O C _ E X C E L . T C L
#
###############################################################################
#                                                                             #
# Copyright(c) 1999/2000/2001/2002/2003/2004/2005/2006  UGS PLM Solutions     #
# Copyright(c) 2007 ~ 2019,                             SIEMENS PLM Software  #
#                                                                             #
###############################################################################
#
# DESCRIPTION:
#
#   This is the main Tcl script that processes and produces the Shop Doc's
#   outputs for a selected object on any of the ONT views.
#
###############################################################################


## - Windows only - Specify path to your EXCEL executable to view the output
##
##   NX will display the result (C:\temp\ug_browser.htm) automatically, when "Display Output" is selected on NX/Shopdoc dailog.
##   In addition, user can configure "::execute_file" below with a proper application to display the result again.

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Set next variable to "1" to display result with external application.
#
set __display_output_ext_app 1


if { $__display_output_ext_app && [string match "windows*" $::tcl_platform(platform)] } {

  #---------------------------------------------------------
  # Specify proper location of the application's executable.
  #
  # set ::execute_file {C:\\apps\\MSOffice\\Office12\\EXCEL.EXE}
   set ::execute_file {C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE}

  # When EXCEL is not available, user can mimic a double-click to open the resultant html file.
  #
   if { ![info exists ::execute_file] || ![file exists $::execute_file] } {
     #<Aug-07-2017 gsl> Enable PB_execute to open resultant file.
      set ::execute_file [file join [MOM_ask_env_var UGII_CAM_SHOP_DOC_DIR] excel_templates PB_execute.exe]
      regsub -all {\\} $::execute_file {\\\\} ::execute_file
   }
}


#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Customize and extend this script -
#
# - The user can customize the handlers defined in this script file by
#   sourceing in "shopdoc_user.tcl". (described at the end of this file).
# - User's script file should reside in the directory defined by
#   UGII_CAM_SHOP_DOC_CUSTOM_DIR, or in user's HOME directory.
#
# For example,
#   User may override DOC__patch_oper_tool_data or DOC__enhance_oper_data to
#   further customize or enhance the functionalities defined in these commands.
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



#=============================================================
proc INFO { args } {
#=============================================================
   if { [info exists ::__Debug_Shopdoc] && $::__Debug_Shopdoc } {
      MOM_output_to_listing_device [join $args]
   }
}


#=============================================================
proc EXEC { command_string {__wait 1} } {
#=============================================================
# This command can be used in place of the intrinsic Tcl "exec" command
# of which some problems have been reported under Win64 O/S and multi-core
# processors environment.
#
#
# Input:
#   command_string -- command string
#   __wait         -- optional flag
#                     1 (default)   = Caller will wait until execution is complete.
#                     0 (specified) = Caller will not wait.
#
# Return:
#   Results of execution
#

   global tcl_platform


   if { $__wait } {

      if { [string match "windows" $tcl_platform(platform)] } {

         global env mom_logname

        # Create a temporary file to collect output
         set result_file "$env(TEMP)/${mom_logname}__EXEC_[clock clicks].out"

        # Clean up existing file
         regsub -all {\\} $result_file {/}  result_file
        #regsub -all { }  $result_file {\ } result_file

         if { [file exists "$result_file"] } {
            file delete -force "$result_file"
         }

         set cmd [concat exec $command_string > \"$result_file\"]
         regsub -all {\\} $cmd {\\\\} cmd

         eval $cmd

        # Return results & clean up temporary file
         if { [file exists "$result_file"] } {
            set fid [open "$result_file" r]
            set result [read $fid]
            close $fid

            file delete -force "$result_file"

           return $result
         }

      } else {

         set cmd [concat exec $command_string]

        return [eval $cmd]
      }

   } else {

      if { [string match "windows" $tcl_platform(platform)] } {

         set cmd [concat exec $command_string &]
         regsub -all {\\} $cmd {\\\\} cmd

        return [eval $cmd]

      } else {

        return [exec $command_string &]
      }
   }
}


#=============================================================
proc PAUSE { args } {
#=============================================================
# Revisions:
#-----------
# 05-19-10 gsl - Use EXEC command
#

   global env

   if { [info exists env(PB_SUPPRESS_UGPOST_DEBUG)]  &&  $env(PB_SUPPRESS_UGPOST_DEBUG) == 1 } {
  return
   }


   global gPB

   if { [info exists gPB(PB_disable_MOM_pause)]  &&  $gPB(PB_disable_MOM_pause) == 1 } {
  return
   }


   global tcl_platform

   set cam_aux_dir  [MOM_ask_env_var UGII_CAM_AUXILIARY_DIR]

   if { [string match "*windows*" $tcl_platform(platform)] } {
      set ug_wish "ugwish.exe"
   } else {
      set ug_wish ugwish
   }

   if { [file exists ${cam_aux_dir}$ug_wish]  &&  [file exists ${cam_aux_dir}mom_pause.tcl] } {

      set title ""
      set msg ""

      if { [llength $args] == 1 } {
         set msg [lindex $args 0]
      }

      if { [llength $args] > 1 } {
         set title [lindex $args 0]
         set msg [lindex $args 1]
      }

      set command_string [concat \"${cam_aux_dir}$ug_wish\" \"${cam_aux_dir}mom_pause.tcl\" \"$title\" \"$msg\"]

      set res [EXEC $command_string]


      switch [string trim $res] {
         no {
            set gPB(PB_disable_MOM_pause) 1
         }
         cancel {
            set gPB(PB_disable_MOM_pause) 1

            uplevel #0 {
               if { [llength [info commands "MOM_abort_program"]] } {
                  MOM_abort_program "*** User Abort Post Processing *** "
               } else {
                  MOM_abort "*** User Abort Post Processing *** "
               }
            }
         }
         default {
            return 
         }
      }

   } else {

      MOM_output_to_listing_device "PAUSE not executed -- \"$ug_wish\" or \"mom_pause.tcl\" not found"
   }
}


#======================================================================================================================================================================================
proc ARR_sort_array_to_list { ARR {by_value 0} {by_decr 0} } {
#=============================================================
# This command will sort and build a list of elements of an array.
#
#   ARR      : Array Name
#   by_value : 0 Sort elements by indices (names - default)
#              1 Sort elements by values
#   by_decr  : 0 Sort into increasing order (default)
#              1 Sort into decreasing order
#
#   Return a list of {name value} couplets
#
#-------------------------------------------------------------
# Feb-24-2016 gsl - Added by_decr flag
#
   upvar $ARR arr

   if { ![info exists arr] } {
return
   }

   set list [list]
   foreach { e v } [array get arr] {
      lappend list "$e $v"
   }

   set val [lindex [lindex $list 0] $by_value]

   if { $by_decr } {
      set decr "decreasing"
   } else {
      set decr "increasing"
   }

   if [expr $::tcl_version > 8.0] {
      if [string is integer "$val"] {
         set list [lsort -integer    -$decr -index $by_value $list]
      } elseif [string is double "$val"] {
         set list [lsort -real       -$decr -index $by_value $list]
      } else {
         set list [lsort -dictionary -$decr -index $by_value $list]
      }
   } else {
      set list [lsort -dictionary -$decr -index $by_value $list]
   }

   foreach v $list {
      append result "$v\n"
   }

return $result
}


#=============================================================
proc Get_NativeName { path } {
#=============================================================
   global tcl_version
   global tcl_platform

   if { [string compare $tcl_version 8.0] >= 0 } {
      return [file nativename $path]
   } else {
      if { [string compare "windows" $tcl_platform(platform)] == 0 } {
        # actually useless.
         regsub -all {/} $path {\\} path
      } elseif { [string compare "unix" $tcl_platform(platform)] == 0 } {
         regsub -all {\\} $path {/} path
      }
   }

return $path
}


#=============================================================
proc DOC__patch_oper_tool_data { } {
#=============================================================
# This command is called in MOM_TOOL_BODY to allow the users
# to enhance the operation list containing parameters of tool
# that would be only available when cycling tool objects.
#
# ==> User may define own version of this command in "shopdoc_user.tcl" (below)
#     to prevent custom data & configuration from being overridden by NX releases.
#
#<Oct-03-2014 gsl> Initial version
#<Feb-24-2015 gsl> Renamed from DOC_patch_oper_tool_data.
#

#<04-06-2015 gsl>
# return


   global mom_tool_name
   global mom_operation_name_data

   global mom_current_ont_view
   global mom_selected_object_name_array
   global mom_selected_object_type_array

  #---------------------------------------
  # Only do this on non-Machine Tool view
  #---------------------------------------
   if [string compare $mom_current_ont_view "MACHVIEW"] {

      global oper_var_list

      if [info exists oper_var_list] {  ;# Variables found on oper list table of the template.

         foreach oper_var $oper_var_list {

            foreach { idx oper } [array get mom_selected_object_name_array] {

               if { $mom_selected_object_type_array($idx) == 100 } { ;# OPER

                 # Oper that uses this tool -
                 #
                  if { [info exists mom_operation_name_data($oper,mom_oper_tool)] &&\
                      ![string compare $mom_operation_name_data($oper,mom_oper_tool) $mom_tool_name] } {

                    #<Apr-11-2016 gsl> Only override undefined vars of "--"
                     if { [info exists mom_operation_name_data($oper,$oper_var)] &&\
                          $mom_operation_name_data($oper,$oper_var) == "--" } {

                       #----------------------------------------------------------------------------
                       # Variables already populated into the data array should have been unset;
                       # here we will only patch up the variables not yet realized within OPER_BODY.
                       #
                        if { [info exists ::$oper_var] } {
                           INFO "Patch up mom_operation_name_data($oper,$oper_var): [set ::$oper_var]"
                           set mom_operation_name_data($oper,$oper_var) [DOC_format_var_with_style $oper_var]

                           unset ::$oper_var
                        }
                     }

                     break
                  }
               }
            }
         }
      }
   }
}
 

#-------------------------------------------------------------
proc MOM_info_user_defined_event { } {
#-------------------------------------------------------------
# This is the callback to be triggered when MOM_list_user_defined_events is called.

   if [info exists ::mom_ude_command_string] {
      INFO "UDE : $::mom_ude_command_string"
   }
}


#=============================================================
proc DOC__enhance_oper_data { } {
#=============================================================
# This command is called in MOM_OPER_BODY to allow the users
# to enhance the operation list containing parameters of toolpath.
# The data would be only available when posting the operation.
#
#   Syntax:
#     MOM_list_oper_path <opr_name> <event_handler> <definition> <output>
#
# ==> User may define own version of this command in "shopdoc_user.tcl" (sourced below)
#     to prevent the customization from being overridden by future NX releases.
#
#-------------------------------------------------------------
#<Feb-24-2015 gsl> Initial version
#

   global mom_operation_name


#<May-16-2018 gsl> Exp. loading data of UDEs - Should be functional in nx12.02.
if 0 {
   MOM_load_oper_ude_exps start $mom_operation_name

   global mom_ude_params_arr
   if [info exists mom_ude_params_arr] {
      INFO "Star UDE of $mom_operation_name \n[array get mom_ude_params_arr]"
   }

   MOM_load_oper_ude_exps end $mom_operation_name

   if [info exists mom_ude_params_arr] {
      INFO "End UDE of $mom_operation_name \n[array get mom_ude_params_arr]"
   }
}

#   MOM_list_user_defined_events start $mom_operation_name


#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# UnComment out next statement to disable subsequent functionalities -
# return


  #<Nov-30-2018 gsl> Devize this mechanism to phish data from tool path of operation -

  # Check if the user has supplied the post for posting -
   set USER_SHOPDOC_DIR [MOM_ask_env_var UGII_CAM_SHOP_DOC_CUSTOM_DIR]
   if { $USER_SHOPDOC_DIR == "" } {
      set USER_SHOPDOC_DIR [MOM_ask_env_var HOME]
   }

   if { $USER_SHOPDOC_DIR != "" } {
      set shopdoc_user_post [file join $USER_SHOPDOC_DIR shopdoc_user_post.tcl]
   }


   if { [info exists shopdoc_user_post] && [file exists $shopdoc_user_post] && [file size $shopdoc_user_post] } {
      set tcl_file [file join $USER_SHOPDOC_DIR shopdoc_user_post.tcl]
      set def_file [file join $USER_SHOPDOC_DIR shopdoc_user_post.def]
   } else {
     # Find sub-post from the same directory as this script -
      set THIS_SHOPDOC_DIR [file dirname $::mom_event_handler_file_name]
      set tcl_file [file join $THIS_SHOPDOC_DIR shopdoc_post.tcl]
      set def_file [file join $THIS_SHOPDOC_DIR shopdoc_post.def]
   }

  # Validate the sub-post and then run it.
   if { [file exists $tcl_file] && [file exists $def_file] } {

     # Create a temp file for communicating data (I/O)
     # => User may also use a fixed file name for this purpose.
      set out_file "${::mom_output_file_directory}${::mom_logname}__tmp_output_[clock clicks].out"

     # With current implementation, MOM_list_oper_path does not return any value!
      MOM_list_oper_path $::mom_operation_name "$tcl_file"\
                                               "$def_file"\
                                               "$out_file"
   } else {
return
   }

  # Process result in the output file produced in the posting job above.
   if { [file exists $out_file] } {

     # Do whatever is needed with the output file, then delete it when so desired.
     # => Data must be retained as global "mom_" variables. They will get picked up
     #    by "::mom_operation_name_data" in MOM_OPER_BODY and presented subsequently.

     if 0 { ;# See below, out_file will be sourced in.
      if { ![catch { set src [open "$out_file" RDONLY] } ] } {
         while { [eof $src] == 0 } {
            set line [gets $src]

           # When data has been prepared & written out in the form of Tcl statements,
           # each line can be evaluated individually here or the output file can be sourced in in its entirety later.
           # => We chose to do the later.
           # eval $line
         }
         close $src
      } else {
        # Error, for some reason output file cannot be open -
      }
     }

     # When the data has been prepared in terms of Tcl syntax, entire output file can be sourced in in one shot -
      if [file size $out_file] {
         source "$out_file"
      }

      file delete "$out_file"

   } else {
     # Handle error if needed -
   }
}


#=============================================================
proc MOM_QUERY_FAIL { } {
#=============================================================
# This handler is triggered when an object does not satisfy the query (QRY) condition.
# It allows some variables to be unset.
#

   #<04-06-2015 gsl> 7110623 - Collect tool data when a single oper is selected on Machine-Tool view
   #                           to produce tool list.
   #                         ==> This should be done here first, since code below will unset variables!
    global mom_selected_object_type_array mom_object_type_name mom_object_type mom_current_ont_view

   # Use this var to register any object (name) being collected to minimize redundancy
    global ex_doc_collected_object_array

    if { [string match "MACHVIEW" $mom_current_ont_view] } {
       if { [array size mom_selected_object_type_array] == 1 && ( $mom_selected_object_type_array(0) == 100 ) } {
          if { [string match "TOOL" $mom_object_type_name] } {

             if { ![info exists ex_doc_collected_object_array($::mom_selected_object_name_array(0))] } {
                set ex_doc_collected_object_array($::mom_selected_object_name_array(0)) 1
                MOM_TOOL_BODY
              return
             }
          }
       }
    }


   #<03-25-11 lxy> unset all the related variables
    global tool_var_list_1
    foreach tool_var $tool_var_list_1 {
      # global $tool_var
       if { [info exists ::$tool_var] && [string compare "mom_group_name" $tool_var] } {
          unset ::$tool_var
       }
    }

   #<03-25-11 lxy> unset all the related variables
    global oper_var_list
    foreach oper_var $oper_var_list {
      # global $oper_var
       if { [info exists ::$oper_var] && [string compare "mom_group_name" $oper_var] } {
          unset ::$oper_var
       }
    }
}


#=============================================================
proc VNC_ask_shared_library_suffix { } {
#=============================================================
   global tcl_platform

   set suffix ""
   set suffix [string trimleft [info sharedlibextension] .]

   if { [string match "" $suffix] } {

      if { [string match "*windows*" $tcl_platform(platform)] } {

         set suffix dll

      } else {

         if { [string match "*HP-UX*" $tcl_platform(os)] } {
            set suffix sl
         } elseif { [string match "*AIX*" $tcl_platform(os)] } {
            set suffix a
         } else {
            set suffix so
         }
      }
   }

return $suffix
}



set mom_operation_name_list [list]
set mom_tool_name_list      [list]
set mom_tool_number_list    [list]


#=============================================================
proc DOC_get_page_height { page_length page_unit PAGE_HEIGHT } {
#=============================================================
#<01-14-11 lxy> change the mechanism of handling page height.
#<01-14-11 gsl> capture page length unit in global

   upvar $PAGE_HEIGHT page_height
   global ex_doc_page_lenght_unit ex_doc_page_lenght_factor

   set ex_doc_page_lenght_unit $page_unit

   if { ![string compare "&nbsp;" $page_length] } {
      set page_length 0
   }

   if { [string compare "IN" $page_unit] && [string compare "MM" $page_unit] } {
      set page_unit "IN"
   }

   if { ![string compare "IN" $page_unit] } {
      set page_height [expr int($page_length * 72.0)]
      set ex_doc_page_lenght_factor 72.0
   } else {
      set page_height [expr int($page_length * 72.0 / 25.4)]
      set ex_doc_page_lenght_factor [expr 72.0/25.4]
   }

}


#=============================================================
proc DOC_Start_Part_Documentation { } {
#=============================================================
    global mom_shop_doc_template_file
    global ex_doc_template_file

    global mom_output_file_directory
    global mom_output_file_basename
    global mom_event_handler_file_name

    global mom_sys_output_file_suffix


   #<Feb-25-2016 gsl> Error encountered, output dir may not be set.
   #                  Redefine dir here early for subsequent consumption.
   # Set working directory to prevent error with unknown "./"

   #<Aug-08-2017 gsl> Resolve file extension from dialog's name entry
    global mom_output_file_suffix

    if { [info exists mom_output_file_basename] } {
       set mom_output_file_basename [string trim $mom_output_file_basename]
       if { [info exists mom_output_file_suffix] && [string trim $mom_output_file_suffix] != "" && \
           ![string match "html" $mom_output_file_suffix] && ![string match "htm" $mom_output_file_suffix] } {
          set mom_output_file_basename ${mom_output_file_basename}.${mom_output_file_suffix}
       }
       set mom_output_file_basename [string trimright ${mom_output_file_basename} .]
       if { [string match "*.html" $mom_output_file_basename] || [string match "*.htm" $mom_output_file_basename] } {
          set mom_sys_output_file_suffix [string trimleft [file extension $mom_output_file_basename] .]
          set mom_output_file_suffix $mom_sys_output_file_suffix
          set mom_output_file_basename [string trim [file rootname $mom_output_file_basename]]
       }
    }

    if { ![info exists mom_sys_output_file_suffix] } {
       if { ![info exists mom_output_file_suffix] || [string trim $mom_output_file_suffix] == "" } {
          set mom_output_file_suffix "html"
       }
       if { ![string match "htm*" $mom_output_file_suffix] } {
          set mom_output_file_suffix "html"
       }
       set mom_sys_output_file_suffix $mom_output_file_suffix
    }

   #<Aug-08-2017 gsl> Write permission to a folder may not be detected properly with the logic above.
    global html_output_file_id
    set html_output_file_id ""

    cd ${mom_output_file_directory}

    MOM_Start_Part_Documentation

    set output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}

    if [catch { set html_output_file_id [open $output_file w] }] {

       cd "[MOM_ask_env_var UGII_TMP_DIR]"

      # Direct outputs to temp dir
       if { [string match "windows" $::tcl_platform(platform)] } {
          set mom_output_file_directory "[file nativename [pwd]]\\"
       } else {
          set mom_output_file_directory "[file nativename [pwd]]/"
       }

       MOM_Start_Part_Documentation

       set output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}

       if [catch { set html_output_file_id [open $output_file w] } res] {
          MOM_abort "Shop Doc output file \"$output_file\" error: $res"
       }
    }

   #<03-18-11 lxy> Refresh the display for getting the PART_GIF image without tool paths.
    catch { MOM_refresh_display }

    if { [info exists mom_shop_doc_template_file] && [file exists $mom_shop_doc_template_file] } {
       set ex_doc_template_file  "$mom_shop_doc_template_file"
    }

    set output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}

   #<03-18-11 lxy> Refresh the display for getting the PART_GIF image without tool paths.
    catch { MOM_refresh_display }

   #<04-18-13 gsl> Add default encoding (for Tcl8.4 & up)
    if { [string compare $::tcl_version "8.0"] > 0 } {
       encoding system "utf-8"
    }

    fconfigure $html_output_file_id -encoding "utf-8"


    if { [llength [info commands "MOM_get_image"]] == 0 } {
       if [file exists [Get_NativeName [file dirname $mom_event_handler_file_name]]/get_image.dll] {
          MOM_run_user_function \
              [Get_NativeName [file dirname $mom_event_handler_file_name]]/get_image.dll ufusr
       }
    }

   #<04-06-11 lxy> the newly created folder name can be defined separately.
    global ex_doc_new_folder_name
    global ex_doc_output_file_structure
    global ex_doc_output_dir
    set ex_doc_output_dir ""

    if { ![info exists ex_doc_new_folder_name] } {
       set ex_doc_new_folder_name ""
    }

    if { ![info exists ex_doc_output_file_structure] } {
       set ex_doc_output_file_structure 0
    }

    if { $ex_doc_output_file_structure } {
       if { ![string compare "" $ex_doc_new_folder_name] } {
          set ex_doc_output_dir "${mom_output_file_directory}${mom_output_file_basename}_files"
          if { $ex_doc_output_file_structure == 1 } {
             set ex_doc_new_folder_name "${mom_output_file_basename}_files/"
          }
       } else {
          set ex_doc_output_dir "${mom_output_file_directory}$ex_doc_new_folder_name"
          if { $ex_doc_output_file_structure == 1 } {
             append ex_doc_new_folder_name "/"
          } else {
             set ex_doc_new_folder_name ""
          }
       }
    }
}


#=============================================================
proc DOC_format_var_with_style { gvar_name } {
#=============================================================
# This function formulates the value of a variable in the style
# specified with the cell of the Excel template.
# ==> Currently, it only handles integers, floating points &
#     scientific notation with given number of decimal places.
#     The 1000th marker(,) will not be processed.
#
# Input:  Name of a global variable (presumably, mom_xxx)
#
# Return: Formatted value if style is found
#
#-------------------------------------------------------------
#<03-27-2013 gsl> Initial version
#<10-15-2013 gsl> No need to format a string
#<08-22-2014 gsl> An array element may come in as gvar_name
#<04-13-2016 gsl> NX11.0
#<01-05-2018 gsl> NX12.02 - Due to change of scheme, this func is disabled.
#-------------------------------------------------------------

   if [info exists ::$gvar_name] {
      set value [set ::$gvar_name]
   } else {
      INFO "Global var ${gvar_name} doesn't exist!"
      upvar $gvar_name value
   }

   if ![info exists value] {
      set value ""
   }

   set result ""
   foreach v [split $value \n] {
      if { $result == "" } {
         set result $v
      } else {
         append result "\n$v"
      }
   }


#<Jan-04-2018 gsl> Short circuit this func, due to different implementation - Retain raw data at this point.
 return $result



   global ex_doc_style_of_var

   set style ""

  #<Jan-03-2018 gsl> Same var can be used in multiple cells rendered with different styles.
  if 1 {
   if [info exists ex_doc_style_of_var($gvar_name)] {
      set style $ex_doc_style_of_var($gvar_name)
      unset ex_doc_style_of_var($gvar_name)
   } else {
      if { [array get ex_doc_style_of_var "$gvar_name,*"] > 0 } {
         set no_style 1
         set cix 0
         while { $no_style } {
            incr cix
            if { [info exists ex_doc_style_of_var($gvar_name,$cix)] } {
               set style $ex_doc_style_of_var($gvar_name,$cix)
               unset ex_doc_style_of_var($gvar_name,$cix)
               set no_style 0
            }
         }
      }
   }
  }

   if { [string length $style] == 0 } {
return $value
   }

   set fmt [DOC_ask_format_of_style $style]
   if { $fmt == "" } {
return $value
   }

   set result ""
   foreach v [split $value \n] {
     #<10-15-2013 gsl> No need to format a string ("string is" may be used in Tcl 8.4 & up.)
      if { ![catch { expr $v }] } {
         set v [format "$fmt" [expr $v]]
      }

      if { $result == "" } {
         set result $v
      } else {
         append result "\n$v"
      }
   }

return $result
}


#=============================================================
proc DOC_ask_format_of_style { style } {
#=============================================================
#
# Input:  Style class
#
# Return: Format string
#
#-------------------------------------------------------------
#<03-29-2013 gsl> Initial version
#-------------------------------------------------------------

   global ex_doc_style_class

   if { [string length $style] > 0  && \
        [info exists ex_doc_style_class($style,mso-number-format)] } {

      set fmt $ex_doc_style_class($style,mso-number-format)

     #<07-19-2019 gsl> Do not format integer or string
      if { [string match "0" $fmt] } {
return ""
      }

     # Digest style format
     # ==> We ignore the 1000 marker "," for now.
      set iexp [string first E $fmt]  ;# Scientific notation
      set idcp [string first . $fmt]  ;# Floating point

      if { $iexp > 0 } {
        # Trim off Exx word
         set fmt [string range $fmt 0 [expr $iexp - 1]]
      }

     # Num of decimal plcs
      if { $idcp < 0 } {
        # No decimal pt, format it as integer
         set ndcp 0
      } else {
         set ndcp [expr [string length $fmt] - $idcp - 1]
      }

      if { $iexp > 0 } {
        # Scientific notation
         set fmt "%.${ndcp}E"
      } else {
         set fmt "%.${ndcp}f"
      }

   } else {
      set fmt ""
   }
}


#=============================================================
proc DOC_parse_excel_html_template_file { } {
#=============================================================
    global mom_shop_doc_template_file
    global ex_doc_template_file
    global mom_sys_output_file_suffix
    global mom_output_file_directory
    global mom_output_file_basename
    global mom_event_handler_file_name


    if { [info exists mom_shop_doc_template_file] && [file exists $mom_shop_doc_template_file] } {
       set ex_doc_template_file  "$mom_shop_doc_template_file"
    }

    set output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}
    MOM_close_output_file $output_file

   #<Aug-08-2017 gsl> (Note) What is happening below ???
    MOM_close_output_file ${mom_output_file_directory}${mom_output_file_basename}.txt
    file delete -force  ${mom_output_file_directory}${mom_output_file_basename}.txt

    if { [llength [info commands "MOM_get_image"]] == 0 } {
        if [file exists [Get_NativeName [file dirname $mom_event_handler_file_name]]/get_image.dll] {
            MOM_run_user_function \
              [Get_NativeName [file dirname $mom_event_handler_file_name]]/get_image.dll ufusr
        }
    }

    global mom_setup_part_gif_file
    global ex_doc_table_index
    global ex_doc_output_title_for_each_page
    global ex_doc_page_height
    global ex_doc_org_file
    global ex_doc_file_bf_body_str
    global ex_doc_file_bf_title_str
    global ex_doc_file_bf_table_str
    global ex_doc_file_of_table
    global ex_doc_file_af_table_end
    global ex_doc_file_af_body_end
    global ex_doc_page_height_bf_title
    global ex_doc_page_height_bf_table
    global ex_doc_table_height
    global ex_doc_page_height_af_table
    global ex_doc_table_title
    global ex_doc_table_line
    global ex_doc_output_dir
    global ex_doc_new_folder_name

    global ex_doc_need_capture_path_gif

   # Open user's template file
    if [catch { set f_id [open "$ex_doc_template_file" r] } res] {
      # Close output file before abort
       global html_output_file_id
       catch { close $html_output_file_id }

       MOM_abort "$ex_doc_template_file can not be open!"
    }

    set org_pic_folder "[file rootname $ex_doc_template_file]_files"
    set org_pic_foler_name "[file tail [file rootname $ex_doc_template_file]]_files"

   # Initialize variables
    array set ex_doc_org_file          [list] ; # content of org file
    array set ex_doc_file_bf_body_str  [list] ; # content before "BODY_START"
    array set ex_doc_file_bf_title_str [list] ; # content before "TITLE_START"
    array set ex_doc_file_bf_table_str [list] ; # content before "TABLE_START"
    array set ex_doc_file_of_table     [list] ; # content of "TABLE"
    array set ex_doc_file_af_table_end [list] ; # content after "TABLE_END"
    array set ex_doc_file_af_body_end  [list] ; # content after "BODY_END"

    set tr_left_num     0 ; # number of "<tr>"
    set tr_right_num    0 ; # number of "</tr>"

    set tr_start       -1 ; # line number of first line of each <tr>...</tr> branch
    set tr_end         -1 ; # line number of last line of each <tr>...</tr> branch
    set body_start     -1 ; # line number of start line of body codes ==> "###BODY_START###"
    set title_start    -1 ; # line number of start line of title codes ==> "###TITLE_START###"
    set table_start    -1 ; # line number of start line of table codes ==> "###TABLE_START###"
    set table_end      -1 ; # line number of end line of table codes ==> "###TABLE_END###"
    set body_end       -1 ; # line number of end line of body codes ==> "###BODY_END###"

    set page_length    ""
    set is_image_br     0

    set is_table_sep    0
    set is_table_title  1
    set is_table_line   0
    set ex_doc_table_index     0

    set ex_doc_page_height_bf_title  0
    set ex_doc_page_height_bf_table  0
    set ex_doc_table_height          0
    set ex_doc_page_height_af_table  0
    set abandon_tr 0
    set image_tr_line ""

    set table_title_str    -1
    set table_title_height  0
    set decrease_key_line_height "height:3pt"

   #<01-19-11 lxy> Space between different tables
    set line_height_between_table "height:20pt"

   # Vars to collect style classes information
    global ex_doc_style_class
    global ex_doc_style_of_var

    set processed_styles [list]

   #++++++++++++++++++++++++++++++++++
   # Start parsing Html template file
   #
    set style_class_begin 0  ;# To signal the encounter of style classes

   # Line counter in Html file
    set i 1

    while { [gets $f_id line] >= 0 } {
       set ex_doc_org_file($i) $line

       if { $body_start == -1 || ![string compare "true" $body_start] } {
         # codes before "###BODY_START###"
          set ex_doc_file_bf_body_str($i) $line

         #---------------------------------------------------
         #<03-27-2013 gsl> Collect style classes information
         #
         if 1 { ;# Setting it to "0" will disable the process for style

          set tline [string trim $line]

         #<Apr-13-2016 gsl> Found case that "style" doesn't begin with </style>
          if { [string match "*<\/style>*" $tline] ||\
               [string match "*<style*" $tline] } {
             if { !$style_class_begin } {
                set style_class_begin 1
                set class ""
             } else {
                set style_class_begin 0
             }
          }
          if { $style_class_begin } {
             if { [string match ".x*" $tline] } {
                set class [string trimleft $tline .]
             }

             if { [string length $class] } {
               # Skip style classes that have been processed,
                if { [lsearch $processed_styles $class] >= 0 } {
                   set class ""
                }
             }

             if { [string length $class] } {

               # Look for format attribute,
                if { [string match "mso-number-format:*" $tline] } {

                  # Grab everything after "mso-number-format:"
                  # ==> Do "eval" to remove extra "\" from the string
                   set fmt [eval string trimright [string range $tline 18 end] \;]

                  # Convert some special formats (in name string)
                  # ==> Not sure if this works for non-English env.!!!
                   if { [string compare "Fixed" $fmt] == 0 } {
                      set fmt "0.00"
                   } elseif { [string compare "Standard" $fmt] == 0 } {
                      set fmt "#,##0.00"
                   } elseif { [string compare "Scientific" $fmt] == 0 } {
                      set fmt "0.00E+00"
                   }

                  # Only store away numeric format (supposedly, containing at least one "0")
                   if [string match "*0*" $fmt] {
                      set ex_doc_style_class($class,mso-number-format) "$fmt"
                   }

                  # Catalog and stop collecting data for this class
                   lappend processed_styles $class
                   set class ""
                }
             }
          }
         } ;# if
         #---------------------------------------------------

       } elseif { $title_start == -1 || ![string compare "true" $title_start] } {
         # codes between "###BODY_START###" and "###TITLE_START###"
          set ex_doc_file_bf_title_str($i) $line

       } elseif { $table_start == -1 || ![string compare "true" $table_start] } {
         # codes between "###TITLE_START###" and "###TABLE_START###"

         # ==> It might be too early to substitute some vars that are not yet available at this point.
         # Replace the variables with their values, if they exist.
         # DOC_subst_exist_var line

          set ex_doc_file_bf_table_str($i) $line

         #<11-apr-2019 gsl> 9366670 - Collect mom_vars used in title block. Info will be consumed by MOM_MEMBERS_FTR.
         # Actually, we only need to know if any mom var is in use to trigger the substituttion.
          if { ![info exists ::ex_doc_title_vars] } {
             set ::ex_doc_title_vars [list]
          }
          set ids [string first "mom_" $line]
          while { $ids >= 0 } {
             set ide [string wordend $line $ids]
             set mom_var [string range $line $ids [expr $ide-1]]
             lappend ::ex_doc_title_vars $mom_var
             set ids [string first "mom_" $line $ide]
          }

       } elseif { $table_end == -1 || ![string compare "true" $table_end] } {
         # codes between "###TABLE_START###" and "###TABLE_END###"
          if { [string match "*###TABLE_SEPARATOR###*" $line] } {
             regsub "###TABLE_SEPARATOR###" $line "" ex_doc_file_of_table($i)
          } else {
             set ex_doc_file_of_table($i) $line
          }

       } elseif { $body_end == -1 || ![string compare "true" $body_end] } {
         # codes between "###TABLE_END###" and "###BODY_END###"
         # Replace the variables with their values, if they exist.
          DOC_subst_exist_var line
          set ex_doc_file_af_table_end($i) $line

       } else {
         # codes after "###BODY_END###"
          set ex_doc_file_af_body_end($i) $line
       }

       set tmp_line [string trim $line]

      # Supposedly, there is only one <tr> (table row) or </tr> in a line. may be problem!!!!!!!
      # find keywords "###*###"
       if { $tr_start >= 0 } {
          if { [string match "*<tr*" $tmp_line] } {
             incr tr_left_num
          } elseif { [string match "*</tr*" $tmp_line] } {
             incr tr_right_num
          }
          if { [string match "*###PAGE_CONTENT_LENGTH###*" $tmp_line] } {
             set page_length true
             set abandon_tr 1
          } elseif { [string match "*###BODY_START###*" $tmp_line] } {
             set body_start true
             set abandon_tr 1
          } elseif { [string match "*###TITLE_START###*" $tmp_line] } {
             set title_start true
             set abandon_tr 1
          } elseif { [string match "*###TABLE_START###*" $tmp_line] } {
             set table_start true
             set abandon_tr 1
          } elseif { [string match "*###TABLE_END###*" $tmp_line] } {
             set table_end true
             set abandon_tr 1
          } elseif { [string match "*###BODY_END###*" $tmp_line] } {
             set body_end true
             set abandon_tr 1
          }

          if { [string match "*if gte vml 1*" $tmp_line] } {
             set is_image_br 1
          }

          if { [string match "*###TABLE_SEPARATOR###*" $tmp_line] } {
             set is_table_sep 1
          }

          if { $tr_left_num == $tr_right_num } {
             set tr_end $i
          }
       }

      # A table row just encountered,
       if { $tr_start == -1 && [string match "<tr*" $tmp_line] } {
          set tr_start $i
          incr tr_left_num

         # Get height of this tr
          set tr_ind [string first "height:" $tmp_line]
          if { $tr_ind == -1 } {
             set tr_height 0
          } else {
             set tmp_tr_line [string range $tmp_line $tr_ind [string length $tmp_line]]
             set tr_ind [string first "pt" $tmp_tr_line]
             set tr_height [string range $tmp_tr_line 7 [expr $tr_ind - 1]]
          }

         # Get the position of start line of image branch
          if { [string match "*if gte vml 1*" $tmp_line] } {
             set is_image_br 1
          }

         # Get the position of "###TABLE_SEPARATOR###" line
          if { [string match "*###TABLE_SEPARATOR###*" $tmp_line] } {
             set is_table_sep 1
          }
       }

       if { $tr_end >= 0 } {
         # Separate title and content for each table
          if { [string compare "true" $table_start] && $table_start >= 0 && $table_end == -1 } {
             if { !$is_table_sep } {
                if { $is_table_title } {

                   if { $ex_doc_table_index == 0 } {
                      set ex_doc_table_title($ex_doc_table_index,start)  $table_title_str
                      set ex_doc_table_title($ex_doc_table_index,height) [expr $tr_height + $table_title_height]
                   } else {
                      set ex_doc_table_title($ex_doc_table_index,start)  $table_title_str
                      set ex_doc_table_title($ex_doc_table_index,height) [expr $tr_height + $table_title_height]
                   }

                   set ex_doc_table_title($ex_doc_table_index,end)    $tr_end
                }

                # Collecting mom vars of each row of a table
                #
                if { $is_table_line } {

                   set ex_doc_table_line($ex_doc_table_index,start)    $tr_start
                   set ex_doc_table_line($ex_doc_table_index,end)      $tr_end
                   set ex_doc_table_line($ex_doc_table_index,height)   $tr_height
                   set ex_doc_table_line($ex_doc_table_index,var_list) [list]

                  #<03-27-2013 gsl> We can not substitute vars here, since they are not available at this moment.
                  #                 ==> Here, we will identify & capture the style class used for the cell <td>
                  #                 ==> A table cell may be written in 2 lines in the Html file,
                  #                     we will pick up the style info then apply it to the subsequent mom var.

                  #<Apr-13-2016 gsl>
                  # "style" is zero'ed ONLY when consumed. This may handle the situation
                  #  where a row is defined in multiple lines in the template (.htm) file

                  # set style ""

                   for { set table_line_ind $tr_start } { $table_line_ind <= $tr_end } { incr table_line_ind } {

                      set tline [string trimright $ex_doc_file_of_table($table_line_ind)]

                     # A typical line would look like "<td class=xl10010074 width=70 ..."
                     # The style class can be identified via "class=" (6 chars) token,
                     # followed by "xl10010074", i.e., will be the class.
                     #
                      set ids [string first "class=" $tline]
                      if { $ids >= 0 } {
                         set ide [string wordend $tline [expr $ids + 6]]
                         set style [string range $tline [expr $ids + 6] [expr $ide - 1]]
                      }

                     # A cell containing a formula would look like:
                     # "... 9pt'>=2*${mom_toolpath_time}+${mom_tool_change_time}</td>",
                     # we may extract the entire expression from the ending </td> to the preceding ">" mark.
                     # ==> A cell may contain an expression with more than 1 mom-var.
                     #
                     # <26-Aug-2016 shuai> Fix PR7781511
                     #                     In an expression, all mom_vars should be added into ex_doc_table_line(*,var_list) and save their values.
                     #                     NOT just the first mom_var.
                     #                     Skip over the current "\$\{mom_" (6 chars) and look for the next one if it exists.
                     if { [info exists style] } {

                       #<Jan-04-2018 gsl> Retain style to be saved for a table data cell
                        set sty $style

                        for { set start_index 0 } { $start_index <= [string length $tline] } { set start_index [expr $mom_ind+6] } {

                          #<Jan-04-2018 gsl>
                           if { $start_index == 0 } { set mom_var "" }

                          # Current scheme won't phish out global var in "::mom_".
                           set mom_ind [string first "\$\{mom_" $tline $start_index]

                           if { $mom_ind >= 0 } {

                              set mom_ind_line [string range $tline $mom_ind [string length $tline]]
                              set first_r_brace [expr [string first "\}" $mom_ind_line] - 1]

                              set mom_var [string range $mom_ind_line 2 $first_r_brace]

                              lappend ex_doc_table_line($ex_doc_table_index,var_list) $mom_var

                             #<Jan-03-2018 gsl> 9043366 - This scheme would always use the last format associate with the var.
                             #                            Problem arises when same var is used in multiple cells.
                             # ==> New scheme does not use style per var -
                             if 0 {
                             # Remember the style for a var
                              if { ![info exists ex_doc_style_of_var($mom_var)] } {
                                 set cix 0
                                 set ex_doc_style_of_var($mom_var) $style
                              } else {
                                # It's assumed the order of styles for a same var collected here will be how the var of cells are rendered.
                                 set no_style 1
                                 set cix 0
                                 while { $no_style } {
                                    incr cix
                                    if { ![info exists ex_doc_style_of_var($mom_var,$cix)] } {
                                       set ex_doc_style_of_var($mom_var,$cix) $style
                                       set no_style 0
                                    }
                                 }
                              }
                             }

                           } else {

                              break
                           }
                        }
                     }

                     # Extract entire expression
                     #
                     if 1 { ;# Setting it to "0" will disable the process for expression

                        set exp ""
                        set ids -1

                       # Trim off trailing </td> (5 chars) ==> What if the line is broken up?
                        if [string match "*</td>" $tline] {
                           set tline [string range $tline 0 [expr [string length $tline] - 6]]

                          # Find last ">"
                           set ids [string last ">" $tline]
                           if { $ids > 0 } {
                             # Line contains ">"; it means entire expression is inline
                              set exp [string range $tline [expr $ids+1] end]
                           } else {
                             # Combine parts to form the expression
                              set exp ${exp_part}${tline}
                           }

                        } else {

                           set exp_part ""

                          # Find last ">"
                           set ids [string last ">" $tline]
                           if { $ids > 0 } {
                             # Expression may be partial
                              set exp_part [string range $tline [expr $ids+1] end]
                           }
                        }

                       # An Excel expression should always start with a "=".
                       # - Convert it to an "EXP:" statement inline...
                        if { [string index $exp 0] == "=" } {

                          # Fudge the line containing an expression
                           set tline [string trimright $ex_doc_file_of_table($table_line_ind)]
                           set tts   [string range $tline 0 $ids]
                           set exp   [string trimleft $exp "="]

                          #<Jan-05-2018 gsl> Handle expression as regular cell;
                          #                  => The exp & cell style (class) will be retained for later evaluation.
                          # set tline ${tts}\{EXP:${style}\ ${exp}\}</td>

                           set tline ${tts}\{${exp}\}</td>

                          #<Apr-13-2016 gsl>
                           unset style

                           set ex_doc_file_of_table($table_line_ind) $tline
                        }

                       #<Jan-04-2018 gsl> Retain exp
                        if { $exp != "" && ![string match "*\&*" $exp] } {
                           set ex_doc_file_of_table($table_line_ind,exp) $exp
                           if ![info exists sty] {
                              set sty ""
                           }
                           set ex_doc_file_of_table($table_line_ind,sty) $sty

                           unset sty
                        }
                     }

                   }
                }

                if { $is_table_line } {
                   set is_table_line 0
                }

                if { $is_table_title } {
                   set is_table_title 0
                   set is_table_line 1
                }

             } else {

                set is_table_title 1
                set is_table_line  0
                set is_table_sep   0
                incr ex_doc_table_index
                set table_title_str $tr_start
                set table_title_height $tr_height

                for { set j $tr_start } { $j <= $tr_end } { incr j } {
                  # regsub {height:[0-9.]+pt} $ex_doc_file_of_table($j) $decrease_key_line_height ex_doc_file_of_table($j)
                  #<05-29-2014 gsl> May not rely on the return value of "regsub"
                   if [regsub {height:[0-9.]+pt} $ex_doc_file_of_table($j) $decrease_key_line_height ex_doc_file_of_table($j)] {
                      set ex_doc_table_title($ex_doc_table_index,head_line_num) $j
                      regsub {height:[0-9.]+pt} $ex_doc_file_of_table($j) $line_height_between_table\
                                                                          ex_doc_table_title($ex_doc_table_index,head_line)
                   }
                }
             }
          }

         # Remove signs such as "### * ###"
          if { ![string compare "true" $page_length] } {
            # Get "page format" and "repeat title"
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                if { [string match "*###PAGE_CONTENT_LENGTH###*" $ex_doc_file_bf_body_str($j)] } {
                   set page_len_line $ex_doc_file_bf_body_str([expr $j + 1])
                   set page_length [string range $page_len_line [expr [string first ">" $page_len_line] + 1]\
                                                                [expr [string first "</" $page_len_line] -1]]
                   set page_unit_line $ex_doc_file_bf_body_str([expr $j + 2])
                   set page_unit [string range $page_unit_line [expr [string first ">" $page_unit_line] + 1]\
                                                               [expr [string first "</" $page_unit_line] -1]]
                } elseif { [string match "*###REPEAT_TITLE###*" $ex_doc_file_bf_body_str($j)] } {
                   set head_repeat_line $ex_doc_file_bf_body_str([expr $j + 1])
                   set head_repeat [string range $head_repeat_line [expr [string first ">" $head_repeat_line] + 1]\
                                                                   [expr [string first "</" $head_repeat_line] -1]]
                }
                unset ex_doc_file_bf_body_str($j)
             }

             DOC_get_page_height $page_length $page_unit ex_doc_page_height

             if { ![string compare $head_repeat "Yes"] || ![string compare $head_repeat "1"] } {
                set ex_doc_output_title_for_each_page 1
             } elseif { ![string compare $head_repeat "No"] || ![string compare $head_repeat 0] } {
                set ex_doc_output_title_for_each_page 0
             } else {
                set ex_doc_output_title_for_each_page 0
             }
          }

          if { ![string compare "true" $body_start] } {
             set body_start [expr $tr_end + 1]
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                unset ex_doc_file_bf_body_str($j)
             }
          }

          if { ![string compare "true" $title_start] } {
             set title_start [expr $tr_end + 1]
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                regsub {height:[0-9.]+pt} $ex_doc_file_bf_title_str($j) $decrease_key_line_height ex_doc_file_bf_title_str($j)
                if { [string match "*###TITLE_START###*" $ex_doc_file_bf_title_str($j)] } {
                   regsub "###TITLE_START###" $ex_doc_file_bf_title_str($j) "" ex_doc_file_bf_title_str($j)
                }
                set ex_doc_file_bf_table_str($j) $ex_doc_file_bf_title_str($j)
                unset ex_doc_file_bf_title_str($j)
             }
          }

          if { ![string compare "true" $table_start] } {
             set table_start [expr $tr_end + 1]
             set table_title_str $tr_start
             set table_title_height $tr_height
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                regsub {height:[0-9.]+pt} $ex_doc_file_bf_table_str($j) $decrease_key_line_height ex_doc_file_bf_table_str($j)
                if { [string match "*###TABLE_START###*" $ex_doc_file_bf_table_str($j)] } {
                   regsub "###TABLE_START###" $ex_doc_file_bf_table_str($j) "" ex_doc_file_bf_table_str($j)
                }
                set ex_doc_file_of_table($j) $ex_doc_file_bf_table_str($j)
                unset ex_doc_file_bf_table_str($j)
             }
          }

          if { ![string compare "true" $table_end] } {
             set table_end [expr $tr_end + 1]
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                unset ex_doc_file_of_table($j)
             }
          }

          if { ![string compare "true" $body_end] } {
             set body_end [expr $tr_end + 1]
             for { set j $tr_start } { $j <= $tr_end } { incr j } {
                unset ex_doc_file_af_table_end($j)
             }
          }

         # Handle image data
          if { $is_image_br } {
             if { $body_start > 0 && $title_start < 0 } \
             {
                set search_str $tr_start
                for { set pic_num 1 } { 1 } { incr pic_num } {
                   set gfxdata_str -1
                   set gfxdata_end -1

                   for { set k $search_str } { $k <= $tr_end } { incr k } {
                      if { $gfxdata_str >= 0 } {
                         if { [string first "\"" $ex_doc_file_bf_title_str($k)] >= 0 } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { [string match "*o:gfxdata=*" $ex_doc_file_bf_title_str($k)] } {
                         set gfxdata_str $k
                         set index1 [expr [string first "o:gfxdata=\"" "$ex_doc_file_bf_title_str($k)"] + 11]
                         set cur_tmp_line [string range "$ex_doc_file_bf_title_str($k)" $index1 [string length "$ex_doc_file_bf_title_str($k)"]]
                         if { [string first "\"" $cur_tmp_line] >=0 } {
                            set gfxdata_end $k
                            set ex_doc_file_bf_title_str($k) \
                                 "[string range $ex_doc_file_bf_title_str($k) 0 [expr $index1 - 1]][string range $cur_tmp_line [string first "\"" $cur_tmp_line] [string length $cur_tmp_line]]"
                            break
                         }
                      }
                   }

                   if { $gfxdata_end > $gfxdata_str } {
                      set ex_doc_file_bf_title_str($gfxdata_str) \
                           [string range $ex_doc_file_bf_title_str($gfxdata_str) \
                                0 [expr [string first "o:gfxdata=" "$ex_doc_file_bf_title_str($gfxdata_str)"] + 10]]

                      for { set image_item [expr $gfxdata_str + 1] } { $image_item < $gfxdata_end } { incr image_item } {
                         unset ex_doc_file_bf_title_str($image_item)
                      }

                      set ex_doc_file_bf_title_str($gfxdata_end) \
                           [string range $ex_doc_file_bf_title_str($gfxdata_end) \
                                [string first "\"" $ex_doc_file_bf_title_str($gfxdata_end)] [string length $ex_doc_file_bf_title_str($gfxdata_end)]]

                      append ex_doc_file_bf_title_str($gfxdata_str) $ex_doc_file_bf_title_str($gfxdata_end)
                      unset ex_doc_file_bf_title_str($gfxdata_end)

                      set search_str [expr $gfxdata_end + 1]

                   } else {
                      unset gfxdata_str gfxdata_end
                      break
                   }
                   unset gfxdata_str gfxdata_end
                }
                unset search_str

                for { set k $tr_start } { $k <= $tr_end } { incr k } {
                   if { [info exists ex_doc_file_bf_title_str($k)] && [string match "*${org_pic_foler_name}/*" $ex_doc_file_bf_title_str($k)] } {
                      regsub -all "$org_pic_foler_name/" $ex_doc_file_bf_title_str($k) "$ex_doc_new_folder_name" ex_doc_file_bf_title_str($k)
                   }
                }
                set is_image_br 0

             } elseif { $title_start > 0 && $table_start < 0 } \
             {
               # Remove data of o:gfxdata=""
               # There may be several pics in one row.
                set search_str $tr_start
               #<03-18-11 lxy> Use ex_doc_is_2003_version to handling pictures separately for office 2003.
                set ex_doc_is_2003_version 0
                for { set pic_num 1 } { 1 } { incr pic_num } {
                   set gfxdata_str -1
                   set gfxdata_end -1

                   for { set k $search_str } { $k <= $tr_end } { incr k } {
                      if { $gfxdata_str >= 0 } {
                         if { [string first "\"" $ex_doc_file_bf_table_str($k)] >= 0 } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { [string match "*o:gfxdata=*" $ex_doc_file_bf_table_str($k)] } {
                         set gfxdata_str $k
                         set index1 [expr [string first "o:gfxdata=\"" "$ex_doc_file_bf_table_str($k)"] + 11]
                         set cur_tmp_line [string range "$ex_doc_file_bf_table_str($k)" $index1 [string length "$ex_doc_file_bf_table_str($k)"]]
                         if { [string first "\"" $cur_tmp_line] >=0 } {
                            set gfxdata_end $k
                            set ex_doc_file_bf_table_str($k) \
                                 "[string range $ex_doc_file_bf_table_str($k) 0 [expr $index1 - 1]][string range $cur_tmp_line [string first "\"" $cur_tmp_line] [string length $cur_tmp_line]]"
                            break
                         }
                      }
                   }

                  #<03-18-11 lxy> for 2003 excel
                   if { $gfxdata_end == -1 && $gfxdata_str == -1 && $pic_num == 1 } {
                      set ex_doc_is_2003_version 1
                   }

                   if { $ex_doc_is_2003_version } {
                      set gfxdata_str $search_str

                      for { set k $gfxdata_str } { $k <= $tr_end } { incr k } {
                         if { [info exists ex_doc_file_bf_table_str($k)] && [string match "*if gte vml 1*" $ex_doc_file_bf_table_str($k)] } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { $gfxdata_end == -1 } {
                         unset gfxdata_str gfxdata_end
                         break
                      }
                   }

                   if { $gfxdata_end > $gfxdata_str } {
                      if { !$ex_doc_is_2003_version } {
                         set ex_doc_file_bf_table_str($gfxdata_str) \
                              [string range $ex_doc_file_bf_table_str($gfxdata_str) \
                                   0 [expr [string first "o:gfxdata=" "$ex_doc_file_bf_table_str($gfxdata_str)"] + 10]]

                         for { set image_item [expr $gfxdata_str + 1] } { $image_item < $gfxdata_end } { incr image_item } {
                            unset ex_doc_file_bf_table_str($image_item)
                         }

                         set ex_doc_file_bf_table_str($gfxdata_end) \
                              [string range $ex_doc_file_bf_table_str($gfxdata_end) \
                                   [string first "\"" $ex_doc_file_bf_table_str($gfxdata_end)] [string length $ex_doc_file_bf_table_str($gfxdata_end)]]

                         append ex_doc_file_bf_table_str($gfxdata_str) $ex_doc_file_bf_table_str($gfxdata_end)
                         unset ex_doc_file_bf_table_str($gfxdata_end)
                      }

                     # Find the first <table> after the picture
                      set use_setup_part_gif 0
                      set tab_str -1
                      set tab_end -1
                      for { set k [expr $gfxdata_end + 1] } { $k <= $tr_end } { incr k } {
                         if { $tab_str == -1 && [string match "*src=*.png*" $ex_doc_file_bf_table_str($k)] } {
                            set src_str [string range $ex_doc_file_bf_table_str($k) [string first "src=" $ex_doc_file_bf_table_str($k)] end]
                            set src_str [string range $src_str [expr [string first "/" $src_str] + 1] end]
                            set src_str [string range $src_str 0 [expr [string first "\"" $src_str] - 1]]
                         }

                         if { [string match "*<table*" $ex_doc_file_bf_table_str($k)] } {
                            set tab_str $k
                         }

                         if { $tab_str != -1 && [string match "*</table>*" $ex_doc_file_bf_table_str($k)] } {
                            set tab_end $k
                            break
                         }
                      }
                      for { set k $tab_str } { $k <= $tab_end } { incr k } {
                         if { [string match "*PART_GIF*" $ex_doc_file_bf_table_str($k)] } {
                            set use_setup_part_gif 1
                            break
                         }
                      }

                     #<03-25-11 lxy> Place all the output files in the same folder
                      for { set k [expr $gfxdata_end + 1] } { $k <= $tab_str } { incr k } {
                         if { $use_setup_part_gif } {
                            if { [info exists ex_doc_file_bf_table_str($k)] && [string first "src=" $ex_doc_file_bf_table_str($k)] >= 0 } {
                               regsub -all "src=\"\[^\"\]*\"" $ex_doc_file_bf_table_str($k) "src=\"${ex_doc_new_folder_name}\${ex_doc_gif_name}\"" ex_doc_file_bf_table_str($k)
                            }
                         } else {
                            if { [info exists ex_doc_file_bf_table_str($k)] && [string first "src=" $ex_doc_file_bf_table_str($k)] >= 0 } {
                               regsub -all "src=\"\[^\"\]*/" $ex_doc_file_bf_table_str($k) "src=\"${ex_doc_new_folder_name}" ex_doc_file_bf_table_str($k)
                            }
                            if { [info exists ex_doc_file_bf_table_str($k)] && [string first "src=" $ex_doc_file_bf_table_str($k)] >= 0 } {
                               regsub -all "/*.gif" $ex_doc_file_bf_table_str($k) "/$src_str" ex_doc_file_bf_table_str($k)
                            }
                         }
                      }

                      set search_str [expr $tab_end + 1]
                      unset tab_str tab_end
                   } else {
                      unset gfxdata_str gfxdata_end
                      break
                   }

                   unset gfxdata_str gfxdata_end
                }
                unset search_str

               #<03-18-11 lxy>
                if { [info exists ex_doc_is_2003_version] } {
                   unset ex_doc_is_2003_version
                }

               # Get the new name list of image content in ex_doc_file_bf_table_str
                set image_list [list]
                set image_list [lsort -integer [array names ex_doc_file_bf_table_str]]
                set image_str [lsearch -exact $image_list $tr_start]
                set image_end [lsearch -exact $image_list $tr_end]
                set image_list [lrange $image_list $image_str $image_end]

               # Replace some key words for image
                set pat "src=\"\[^\"\]*\""
                for { set image_it $tr_start } { $image_it <= $tr_end } { incr image_it } {
                   if { [info exists ex_doc_file_bf_table_str($image_it)] } {
                      if { [string first "alt=" $ex_doc_file_bf_table_str($image_it)] >= 0 } {
                         regsub -all "alt=\"\[^\"\]*\"" $ex_doc_file_bf_table_str($image_it) "" ex_doc_file_bf_table_str($image_it)
                      }
                      if { [string first "\[" $ex_doc_file_bf_table_str($image_it)] >= 0 } {
                         regsub -all "\\\[" $ex_doc_file_bf_table_str($image_it) "\\\[" ex_doc_file_bf_table_str($image_it)
                      }
                      if { [string first "\]" $ex_doc_file_bf_table_str($image_it)] >= 0 } {
                         regsub -all "\]" $ex_doc_file_bf_table_str($image_it) "\\\]" ex_doc_file_bf_table_str($image_it)
                      }
                   }
                }
                set is_image_br 0

             } elseif { $table_start > 0 && $table_end < 0 } \
             {
                set search_str $tr_start
               #<03-18-11 lxy> Use ex_doc_is_2003_version to handling pictures separately for office 2003.
                set ex_doc_is_2003_version 0
                for { set pic_num 1 } { 1 } { incr pic_num } {
                   set gfxdata_str -1
                   set gfxdata_end -1
                   set table_end -1
                   for { set k $search_str } { $k <= $tr_end } { incr k } {
                      if { $gfxdata_str >= 0 } {
                         if { [string first "\"" $ex_doc_file_of_table($k)] >= 0 } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { [string match "*o:gfxdata=*" $ex_doc_file_of_table($k)] } {
                         set gfxdata_str $k
                         set index1 [expr [string first "o:gfxdata=\"" "$ex_doc_file_of_table($k)"] + 11]
                         set cur_tmp_line [string range "$ex_doc_file_of_table($k)" $index1 [string length "$ex_doc_file_of_table($k)"]]
                         if { [string first "\"" $cur_tmp_line] >=0 } {
                            set gfxdata_end $k
                            set ex_doc_file_of_table($k) \
                                 "[string range $ex_doc_file_of_table($k) 0 [expr $index1 - 1]][string range $cur_tmp_line [string first "\"" $cur_tmp_line] [string length $cur_tmp_line]]"
                            break
                         }
                      }
                   }

                  #<03-18-11 lxy> For 2003 excel
                   if { $gfxdata_end == -1 && $gfxdata_str == -1 && $pic_num == 1 } {
                      set ex_doc_is_2003_version 1
                   }

                   if { $ex_doc_is_2003_version } {
                      set gfxdata_str $search_str

                      for { set k $gfxdata_str } { $k <= $tr_end } { incr k } {
                         if { [info exists ex_doc_file_of_table($k)] && [string match "*</table>*" $ex_doc_file_of_table($k)] } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { $gfxdata_end == -1 } {
                         unset gfxdata_str gfxdata_end
                         break
                      }
                   }

                   if { $gfxdata_end > $gfxdata_str } {
                      if { !$ex_doc_is_2003_version } {
                         set ex_doc_file_of_table($gfxdata_str) \
                              [string range $ex_doc_file_of_table($gfxdata_str) \
                                   0 [expr [string first "o:gfxdata=" "$ex_doc_file_of_table($gfxdata_str)"] + 10]]

                         for { set image_item [expr $gfxdata_str + 1] } { $image_item < $gfxdata_end } { incr image_item } {
                            unset ex_doc_file_of_table($image_item)
                         }

                         for { set k $gfxdata_str } { $k <= $tr_end } { incr k } {
                            if { [info exists ex_doc_file_of_table($k)] && [string match "*</table>*" $ex_doc_file_of_table($k)] } {
                               set table_end $k
                               break
                            }
                         }

                         set ex_doc_file_of_table($gfxdata_end) \
                              [string range $ex_doc_file_of_table($gfxdata_end) \
                                   [string first "\"" $ex_doc_file_of_table($gfxdata_end)] [string length $ex_doc_file_of_table($gfxdata_end)]]

                         append ex_doc_file_of_table($gfxdata_str) $ex_doc_file_of_table($gfxdata_end)
                         unset ex_doc_file_of_table($gfxdata_end)
                      }

                     # End of the gif code
                      set tmp_gif_end 0
                      for { set k $gfxdata_str } { $k <= $tr_end } { incr k } {
                         if { [info exists ex_doc_file_of_table($k)] && [string match "*</table>*" $ex_doc_file_of_table($k)] } {
                            set tmp_gif_end $k
                            break
                         }
                      }
                      if { !$tmp_gif_end } {
                         set tmp_gif_end $tr_end
                      }

                     # Whether the gif is PATH_GIF
                      set use_path_gif 0
                      for { set k $gfxdata_str } { $k <= $tmp_gif_end } { incr k } {
                         if { [info exists ex_doc_file_of_table($k)] && [string match "*PATH_GIF*" $ex_doc_file_of_table($k)] } {
                            regsub -all "PATH_GIF" $ex_doc_file_of_table($k) "" ex_doc_file_of_table($k)
                            set use_path_gif 1
                            set ex_doc_need_capture_path_gif 1
                            break
                         }
                      }

                     #<03-25-11 lxy> Place all the output files in the same folder
                      for { set k $search_str } { $k <= $tmp_gif_end } { incr k } {

                         if $use_path_gif {
                            if { [info exists ex_doc_file_of_table($k)] && [string first "src=" $ex_doc_file_of_table($k)] >= 0 } {
                               regsub -all "src=\"\[^\"\]*\"" $ex_doc_file_of_table($k) "src=\"${ex_doc_new_folder_name}\${mom_operation_name}_path.gif\"" ex_doc_file_of_table($k)
                            }
                            if { [info exists ex_doc_file_of_table($k)] && [string first "href=" $ex_doc_file_of_table($k)] >= 0 } {
                               regsub -all "href=\"\[^\"\]*\"" $ex_doc_file_of_table($k) "href=\"${ex_doc_new_folder_name}\${mom_operation_name}_path.gif\"" ex_doc_file_of_table($k)
                            }

                         } else {
                            if { [info exists ex_doc_file_of_table($k)] && [string match "*${org_pic_foler_name}/*" $ex_doc_file_of_table($k)] } {
                               regsub -all "$org_pic_foler_name/" $ex_doc_file_of_table($k) "${ex_doc_new_folder_name}" ex_doc_file_of_table($k)
                            }
                         }
                      }
                      set use_path_gif 0


                      set search_str [expr $tmp_gif_end + 1]
                   } else {
                      unset gfxdata_str gfxdata_end
                      break
                   }
                   unset table_end
                   unset gfxdata_str
                   unset gfxdata_end
                }
                unset search_str

               #<03-18-11 lxy> For 2003 excel
                if { [info exists ex_doc_is_2003_version] } {
                   unset ex_doc_is_2003_version
                }

                set is_image_br 0

             } elseif { $table_end > 0 && $body_end < 0 } \
             {

                set search_str $tr_start
                for { set pic_num 1 } { 1 } { incr pic_num } {
                   set gfxdata_str -1
                   set gfxdata_end -1

                   for { set k $search_str } { $k <= $tr_end } { incr k } {
                      if { $gfxdata_str >= 0 } {
                         if { [string first "\"" $ex_doc_file_af_table_end($k)] >= 0 } {
                            set gfxdata_end $k
                            break
                         }
                      }

                      if { [string match "*o:gfxdata=*" $ex_doc_file_af_table_end($k)] } {
                         set gfxdata_str $k
                         set index1 [expr [string first "o:gfxdata=\"" "$ex_doc_file_af_table_end($k)"] + 11]
                         set cur_tmp_line [string range "$ex_doc_file_af_table_end($k)" $index1 [string length "$ex_doc_file_af_table_end($k)"]]
                         if { [string first "\"" $cur_tmp_line] >=0 } {
                            set gfxdata_end $k
                            set ex_doc_file_af_table_end($k) \
                                 "[string range $ex_doc_file_af_table_end($k) 0 [expr $index1 - 1]][string range $cur_tmp_line [string first "\"" $cur_tmp_line] [string length $cur_tmp_line]]"
                            break
                         }
                      }
                   }

                   if { $gfxdata_end > $gfxdata_str } {
                      set ex_doc_file_af_table_end($gfxdata_str) \
                           [string range $ex_doc_file_af_table_end($gfxdata_str) \
                                0 [expr [string first "o:gfxdata=" "$ex_doc_file_af_table_end($gfxdata_str)"] + 10]]

                      for { set image_item [expr $gfxdata_str + 1] } { $image_item < $gfxdata_end } { incr image_item } {
                         unset ex_doc_file_af_table_end($image_item)
                      }

                      set ex_doc_file_af_table_end($gfxdata_end) \
                           [string range $ex_doc_file_af_table_end($gfxdata_end) \
                                [string first "\"" $ex_doc_file_af_table_end($gfxdata_end)] [string length $ex_doc_file_af_table_end($gfxdata_end)]]

                      append ex_doc_file_af_table_end($gfxdata_str) $ex_doc_file_af_table_end($gfxdata_end)
                      unset ex_doc_file_af_table_end($gfxdata_end)

                      set search_str [expr $gfxdata_end + 1]

                   } else {
                      unset gfxdata_str gfxdata_end
                      break
                   }
                   unset gfxdata_str gfxdata_end
                }
                unset search_str

                for { set k $tr_start } { $k <= $tr_end } { incr k } {
                   if { [info exists ex_doc_file_af_table_end($k)] && [string match "*${org_pic_foler_name}/*" $ex_doc_file_af_table_end($k)] } {
                      regsub -all "$org_pic_foler_name/" $ex_doc_file_af_table_end($k) "${ex_doc_new_folder_name}" ex_doc_file_af_table_end($k)
                   }
                }
                set is_image_br 0
             }

          }

         # Get height for each line
          if { $abandon_tr == 0 } {
             if { $title_start == -1 } {
                set ex_doc_page_height_bf_title [expr $ex_doc_page_height_bf_title + $tr_height]
             } elseif { $table_start == -1 } {
                set ex_doc_page_height_bf_table [expr $ex_doc_page_height_bf_table + $tr_height]
             } elseif { $table_end == -1 } {
                set ex_doc_table_height [expr $ex_doc_table_height + $tr_height] ; # should separately save the height of table
             } else {
                set ex_doc_page_height_af_table [expr $ex_doc_page_height_af_table + $tr_height]
             }
          } else {
             set abandon_tr 0
          }

          set tr_left_num     0
          set tr_right_num    0
          set tr_start       -1
          set tr_end         -1
       }

       incr i
    }

   # Collect variables used in template
    global oper_var_list tool_var_list_1

    set oper_var_list   [list]
    set tool_var_list_1 [list]

    foreach var_list_item [array names ex_doc_table_line "*,var_list"] {

       set tmp_var_list $ex_doc_table_line($var_list_item)

      # Identify 1st column (anchor) of the table to be populated
       if { ![string compare "mom_operation_name" [lindex $tmp_var_list 0]] } {
          for { set k 1 } { $k < [llength $tmp_var_list] } { incr k } {
             set tmp_var [lindex $tmp_var_list $k]
             if { [lsearch -exact $oper_var_list $tmp_var] == -1 } {
                lappend oper_var_list $tmp_var
             }
          }
       } elseif { ![string compare "mom_tool_name" [lindex $tmp_var_list 0]] } {
          for { set k 1 } { $k < [llength $tmp_var_list] } { incr k } {
             set tmp_var [lindex $tmp_var_list $k]
             if { [lsearch -exact $tool_var_list_1 $tmp_var] == -1 } {
                lappend tool_var_list_1 $tmp_var
             }
          }
       } elseif { ![string compare "mom_tool_number" [lindex $tmp_var_list 0]] } {
         #<Apr-13-2016 gsl> Also collect tool_name_data
          lappend tool_var_list_1 [lindex $tmp_var_list 0]
          for { set k 1 } { $k < [llength $tmp_var_list] } { incr k } {
             set tmp_var [lindex $tmp_var_list $k]
             if { [lsearch -exact $tool_var_list_1 $tmp_var] == -1 } {
                lappend tool_var_list_1 $tmp_var
             }
          }
       } else {
         # Close output file then abort
          global html_output_file_id
          catch { close $html_output_file_id }

          MOM_abort "The first column (not including Index column) must be \"mom_operation_name\", \"mom_tool_number\" or \"mom_tool_name\"."
       }
    }

   # close user's template file
    close $f_id
}


#=============================================================
proc MOM_Start_Part_Documentation { } {
#=============================================================
   global mom_output_file_directory mom_output_file_basename
   global mom_sys_output_file_suffix

   #<Aug-08-2017 gsl> Interactive mode, this handler is called first
    if { ![info exists mom_sys_output_file_suffix] } {
return
    }

  # set mom_sys_output_file_suffix "html"

  #<07-27-2011 gsl> Prevent NX from crash when the shop-doc output file has been locked somehow.
  #                 Existing file will be overridden.
   if { [file exists ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}] &&\
        [file writable ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}] } {

      MOM_close_output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}

      if [catch { file delete -force ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix} } res] {
        # This msg kind of annoying, hide it for the time being...
        # INFO "$res\nOutput file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix} is locked!"
      }
   }
}


#=============================================================
proc MOM_SETUP_HDR { } {
#=============================================================
    global mom_event_handler_file_name
}


#=============================================================
proc MOM_SETUP_BODY { } {
#=============================================================
    global mom_setup_part_gif_file
}


#=============================================================
proc DOC_prepare_folder { } {
#=============================================================
    global mom_output_file_directory
    global mom_event_handler_file_name
    global mom_output_file_basename
    global mom_setup_part_gif_file
    global ex_doc_gif_name
    global ex_doc_output_dir

   # If there are no data to output, no need to do anything.
    global mom_operation_name_list
    global mom_tool_name_list
    global ex_doc_no_data_generated

    set ex_doc_no_data_generated 0
    if { [llength $mom_operation_name_list] == 0 && [llength $mom_tool_name_list] == 0 } {
       set ex_doc_no_data_generated 1
       return
    }

    if { [file exists $ex_doc_output_dir] } {
        file delete -force $ex_doc_output_dir
    }

   #<04-02-11> Control the structure of the output files.
    global ex_doc_output_file_structure

    if { ![info exists ex_doc_output_file_structure] } {
       set ex_doc_output_file_structure 0
    }
    if { $ex_doc_output_file_structure } {
       file mkdir $ex_doc_output_dir
    }


    set mom_setup_part_gif_file \
        [Get_NativeName ${mom_output_file_directory}${mom_output_file_basename}.gif]

    if { [info commands "MOM_get_image"] != "" } {
       if [catch { MOM_get_image $mom_setup_part_gif_file } res] {
          INFO "$res"
       }
    } else {
      # When SETUP_part.gif exists, just use it.
       set cam_setup_gif [Get_NativeName ${mom_output_file_directory}SETUP_part.gif]

       if { [file exists $cam_setup_gif] } {
          file rename -force $cam_setup_gif $mom_setup_part_gif_file
       } else {
          if [catch { MOM_refresh_display } res] {
             INFO "$res"
          }
          if [catch { MOM_capture_image $mom_setup_part_gif_file } res] {
             INFO "$res"
          }
       }
    }

   #<04-02-11> Control the structure of the output files.
    if { $ex_doc_output_file_structure } {
       if [file exists $mom_setup_part_gif_file] {
          file copy -force $mom_setup_part_gif_file $ex_doc_output_dir
          file delete -force $mom_setup_part_gif_file
       }
    }

    #set cam_setup_gif [Get_NativeName ${mom_output_file_directory}SETUP_part.gif]
    #if [file exists $cam_setup_gif] {
    #   file copy -force $cam_setup_gif $ex_doc_output_dir
    #   file delete -force $cam_setup_gif
    #}

    set ex_doc_gif_name [file tail $mom_setup_part_gif_file]
}


#=============================================================
proc DOC_subst_exist_var { LINE } {
#=============================================================
# Replace the variables with their values
#
# ==> All variables used in the template MUST BE "globals"
#     to get evaluated properly here!!!

   upvar $LINE line

   set tmp_line ""

   set mom_ind_str [string first "\$\{" $line]
   if { $mom_ind_str >= 0 } {

      set mom_ind_line [string range $line $mom_ind_str end]
      set first_r_brace [expr [string first "\}" $mom_ind_line] - 1]

      if { $first_r_brace >= 0 } {
         set mom_ind_end [expr $mom_ind_str + $first_r_brace + 2]

        #<Apr-14-2016 gsl> Somehow mom_var_name may have multiple.
        #                  ==> But it has to be dealt with in "info exists" command; otherwise,
        #                      the actual var (::mom_Output) becomes undefined and gets INFO below.
        # set mom_var_name [lindex [string range $mom_ind_line 2 $first_r_brace] 0]
         set mom_var_name [string range $mom_ind_line 2 $first_r_brace]

         append tmp_line [string range $line 0 [expr $mom_ind_str - 1]]

         if { [info exists ::$mom_var_name] } {

            if { ![string compare $mom_var_name "mom_part_name"] } {
               append tmp_line "[file rootname [file tail [set ::$mom_var_name]]]"
            } else {
               regsub -all {\\} "[set ::$mom_var_name]" {\\\\} tmp_var
               append tmp_line ${tmp_var}
            }
         } else {
            append tmp_line "--"
            INFO "$mom_var_name not found"
         }

         append tmp_line [string range $line $mom_ind_end end]
         set line $tmp_line
      }
   }
}











#=============================================================
proc MOM_End_Part_Documentation { } {
#=============================================================

   global ex_doc_table_index
   global ex_doc_gif_name
   global mom_event_handler_file_name
   global mom_output_file_directory
   global mom_output_file_basename
   global mom_sys_output_file_suffix

   global ex_doc_org_file
   global ex_doc_file_bf_body_str
   global ex_doc_file_bf_title_str
   global ex_doc_file_bf_table_str
   global ex_doc_file_of_table
   global ex_doc_file_af_table_end
   global ex_doc_file_af_body_end
   global ex_doc_page_height_bf_title
   global ex_doc_page_height_bf_table
   global ex_doc_table_height
   global ex_doc_page_height_af_table
   global ex_doc_table_title
   global ex_doc_table_line

   global ex_doc_page_height                 ;# max height of table for each page.
   global ex_doc_output_title_for_each_page  ;# decide whether each page has a title.

   global mom_operation_name_list
   global mom_tool_name_list
   global mom_tool_number_list
   global mom_operation_name_data
   global mom_tool_name_data
   global mom_tool_number_data

   global html_output_file_id

  # If there are no data to output, no need to do anything.
   global ex_doc_no_data_generated
   if { $ex_doc_no_data_generated } {
return
   }

   if { ![info exists ex_doc_page_height] } {
       set ex_doc_page_height 710
   }

   if { ![info exists ex_doc_output_title_for_each_page] } {
       set ex_doc_output_title_for_each_page 0
   }

  #<Apr-14-2016 gsl> Determine object names per selected objects
  # When current view is machine tool, lone tools may have been selected.
  # Tool list will be merged with those used by the opers involved.
  #
   if [string match "MACHVIEW" $::mom_current_ont_view] {
     # When doing tool list
      if { [llength $::tool_var_list_1] > 0 } {
         set tool_list [list]
         foreach { idx obj } [join [split [ARR_sort_array_to_list ::mom_selected_object_name_array] \n]] {
           # Collect tool names
            if { $::mom_selected_object_type_array($idx) == 109 } { ;# TOOL
               lappend tool_list $obj
            }
         }
         if [info exists ::OPER_tool_name_list] {
            foreach t $::OPER_tool_name_list {
               if { [lsearch $tool_list $t] < 0 } {
                  lappend tool_list $t
               }
            }
         }
         if { [llength $tool_list] > 0 } {
            set ::OPER_tool_name_list $tool_list
         }
      }
   }


  #<03-28-2013 gsl> Reformat table data with style
   global oper_var_list

   if [info exists mom_operation_name_data] {
      foreach oper_name $mom_operation_name_list {
         foreach oper_var $oper_var_list {
           # Borrow the value of a var to get formatted
           #<Apr-11-2016 gsl> Direct access to global $oper_var

           #<Jan-03-2018 gsl> Added error protect (as for tool data below)
            if { [info exists mom_operation_name_data($oper_name,$oper_var)] } {
               set ::$oper_var $mom_operation_name_data($oper_name,$oper_var)
               set mom_operation_name_data($oper_name,$oper_var) [DOC_format_var_with_style $oper_var]
            }
         }
      }
   }


   global tool_var_list_1

   if [info exists mom_tool_name_data] {
     #<07-10-2014 gsl> Replace tool list from oper's
      global OPER_tool_name_list
      if [info exists OPER_tool_name_list] {
         set mom_tool_name_list $OPER_tool_name_list
      }

      foreach tool_name $mom_tool_name_list {
         foreach tool_var $tool_var_list_1 {
           # Borrow the value of a var to get formatted
           # global $tool_var <== Do Not decalre any MOM variable global here, since we don't want to alter the actual var in the global scope.

           #<10-21-2014 gsl> Error protect
            if { [info exists mom_tool_name_data($tool_name,$tool_var)] } {
               set ::$tool_var $mom_tool_name_data($tool_name,$tool_var)
               set mom_tool_name_data($tool_name,$tool_var) [DOC_format_var_with_style $tool_var]
            }
         }
      }
   }

   set table_list [array names ex_doc_table_line "*,var_list"]

   foreach table_item $table_list {

      set table_var_list $ex_doc_table_line($table_item)
      set first_var [lindex $table_var_list 0] ;# Fetch 1st var name of the table

     #<Apr-13-2016 gsl> Force tool numbers list to use data gathered for tool names list.
      if [string match "mom_tool_number" $first_var] {
         set first_var "mom_tool_name"
      }

      foreach first_var_it [set ${first_var}_list] { ;# mom_operation_name_list, mom_tool_name_list or mom_tool_number_list

         set SD_data_list($first_var_it) [list]
         foreach s_var [lrange $table_var_list 1 end] {
            if { [info exists ${first_var}_data($first_var_it,$s_var)] } {

               lappend SD_data_list($first_var_it) [set ${first_var}_data($first_var_it,$s_var)]

            } else { ;# When / how would this occur???

               if { [string compare "mom_operation_name" $first_var] == 0 } {

                     if { [llength $mom_tool_name_list] > 0 } {
                        if { [catch { set tool_name $mom_operation_name_data($first_var_it,mom_oper_tool) } res] } {
                           lappend SD_data_list($first_var_it) "\"--\""
                        } else {
                           if { [info exists mom_tool_name_data($tool_name,$s_var)] } {
                              lappend SD_data_list($first_var_it) $mom_tool_name_data($tool_name,$s_var)
                           } else {
                              lappend SD_data_list($first_var_it) "\"--\""
                           }
                        }
                     } else {
                        if { [catch { set tool_number $mom_operation_name_data($first_var_it,mom_tool_number) } res] } {
                           lappend SD_data_list($first_var_it) "\"--\""
                        } else {
                           if { [info exists mom_tool_number_data($tool_number,$s_var)] } {
                              lappend SD_data_list($first_var_it) $mom_tool_number_data($tool_number,$s_var)
                           } else {
                              lappend SD_data_list($first_var_it) "\"--\""
                           }
                        }
                     }

               } elseif { [string compare "mom_tool_name" $first_var] == 0 } {
                    #???
               }
            }
         }
      }

      #<06-29-11 lxy> "--" was "No information"
      if { [llength [set ${first_var}_list]] == 0 } {
         set ${first_var}_list [list "\"--\""]
         set tmp_first_var "\"--\""
         set SD_data_list($tmp_first_var) [list]
         foreach s_var [lrange $table_var_list 1 end] {
            lappend SD_data_list($tmp_first_var) "\"--\""
         }
      }
   }

  # Open result file
   if { $html_output_file_id != "" } {

      set output_file ${mom_output_file_directory}${mom_output_file_basename}.${mom_sys_output_file_suffix}
      set fr_id $html_output_file_id

      foreach item [lsort -integer [array names ex_doc_file_bf_body_str]] {
         puts $fr_id $ex_doc_file_bf_body_str($item)
      }
   } else {
      MOM_abort "Shop Doc output file not found!"
   }

  #<01-14-11 lxy> limit the value of ex_doc_page_height, when it's too small.
   set template_page_height [expr $ex_doc_page_height_bf_title + $ex_doc_page_height_bf_table + $ex_doc_page_height_af_table]

  #<01-14-11 gsl> If title is repeated (min page length = title block + one row) otherwise (= title block).
   if { $ex_doc_output_title_for_each_page } {
      set template_page_height [expr $template_page_height + $ex_doc_table_title(0,height) + $ex_doc_table_line(0,height)]
   }

   if { ($ex_doc_page_height > 0) && ($ex_doc_page_height < $template_page_height) } {
      global ex_doc_page_lenght_unit ex_doc_page_lenght_factor
      INFO "*** Given page length ([expr $ex_doc_page_height/$ex_doc_page_lenght_factor] $ex_doc_page_lenght_unit)\
                should not be less than the minimal length for this template\
                ([expr $template_page_height/$ex_doc_page_lenght_factor] $ex_doc_page_lenght_unit). ***"

      set ex_doc_page_height $template_page_height
   }

   unset template_page_height


  #<01-14-11 lxy> when the value of ex_doc_page_height is 0, output all the content in one page.
   if { $ex_doc_page_height > 0 } {

     # Calculate the height and the number of page
      set allowed_table_height [expr $ex_doc_page_height - $ex_doc_page_height_bf_title -\
                                     $ex_doc_page_height_bf_table - $ex_doc_page_height_af_table]
      set page_number 1
      set tmp_height 0
      array set page [list]

      for { set index 0 } { $index <= $ex_doc_table_index } { incr index } {
         set first_var [lindex $ex_doc_table_line($index,var_list) 0]
         set tmp_height [expr $tmp_height + $ex_doc_table_title($index,height)]

         if { $tmp_height > [expr $allowed_table_height - $ex_doc_table_line($index,height)] } {
            if { [info exists line_index] } {
               set page($page_number,table) [expr $index - 1]
               set page($page_number,line) [expr $line_index -1]
               if { $page_number == 1 && $ex_doc_output_title_for_each_page == 0 } {
                  set allowed_table_height [expr $ex_doc_page_height - $ex_doc_page_height_bf_title - $ex_doc_page_height_af_table]
               }
            } else {
               set page($page_number,table) 0
               set page($page_number,line)  0

              # the height of head is too big, so,only output it on the first page
               set allowed_table_height [expr $ex_doc_page_height - $ex_doc_page_height_bf_title - $ex_doc_page_height_af_table]
            }

            incr page_number
            set tmp_height $ex_doc_table_title($index,height)
         }

        #<Apr-13-2016 gsl> Critical tweak for tool_number list - 
        #                  to set the number of rows (objects) to be tabulated.
         if [string match "mom_tool_number" $first_var] {
            set tabel($index,line_num) [llength [set mom_tool_name_list]]
         } else {
            set tabel($index,line_num) [llength [set ${first_var}_list]]
         }

         for { set line_index 1 } { $line_index <= $tabel($index,line_num) } { incr line_index } {
            set tmp_height [expr $tmp_height + $ex_doc_table_line($index,height)]
            if { $tmp_height == 0 } {
               set tmp_height [expr $tmp_height + $ex_doc_table_title($index,height)]
            }
            if { $tmp_height > [expr $allowed_table_height - $ex_doc_table_line($index,height)] } {
               set page($page_number,table) $index
               set page($page_number,line) $line_index
               if { $page_number == 1 && $ex_doc_output_title_for_each_page == 0 } {
                  set allowed_table_height [expr $ex_doc_page_height - $ex_doc_page_height_bf_title - $ex_doc_page_height_af_table]
               }
               incr page_number
               set tmp_height 0
            } elseif { $index == $ex_doc_table_index && $line_index == $tabel($index,line_num) } {
               set page($page_number,table) $index
               set page($page_number,line) $line_index
            }
         }
      }

      set page_number [llength [array names page "*,table"]]
      set total_pages $page_number

   } else {

      set page_number 1
      set page($page_number,table) $ex_doc_table_index

      for { set p 0 } { $p <= $ex_doc_table_index } { incr p } {

         set first_var [lindex $ex_doc_table_line($p,var_list) 0]

        #<Apr-13-2016 gsl> Critical tweak for tool_number list - 
        #                  to set the number of rows (objects) to be tabulated.
         if [string match "mom_tool_number" $first_var] {
            set tabel($p,line_num) [llength [set mom_tool_name_list]]
         } else {
            set tabel($p,line_num) [llength [set ${first_var}_list]]
         }
      }

      set page($page_number,line) $tabel($ex_doc_table_index,line_num)
      set total_pages $page_number
   }

   set ta_ind 0
   set is_break 0
   set frist_line_num 1

   for { set page_ind 1 } { $page_ind <= $page_number } { incr page_ind } {

      set cur_page $page_ind
      foreach item [lsort -integer [array names ex_doc_file_bf_title_str]] {
         #puts $fr_id [subst $ex_doc_file_bf_title_str($item)]
         set ttmp_line $ex_doc_file_bf_title_str($item)
         regsub -all {\[} $ttmp_line {\\[}  ttmp_line
         regsub -all {\]} $ttmp_line {\\]}  ttmp_line
         set ttmp_line [subst $ttmp_line]
         regsub -all {\\\[}  $ttmp_line {[}  ttmp_line
         regsub -all {\\\]}  $ttmp_line {]}  ttmp_line
         puts $fr_id $ttmp_line
         unset ttmp_line
      }


     # Page title
      if { $page_ind == 1 || $ex_doc_output_title_for_each_page == 1 } {

         foreach item [lsort -integer [array names ex_doc_file_bf_table_str]] {
            set ttmp_line $ex_doc_file_bf_table_str($item)
            DOC_subst_exist_var ttmp_line

            regsub -all {\[} $ttmp_line {\\[}  ttmp_line
            regsub -all {\]} $ttmp_line {\\]}  ttmp_line
            if { [catch { set output_tmp_line "[subst $ttmp_line]"} res] } {
               set output_tmp_line "$ttmp_line"
            }
            regsub -all {\\\\\[} $output_tmp_line {[}  output_tmp_line
            regsub -all {\\\\\]} $output_tmp_line {]}  output_tmp_line

           # ==> Do we need another round of subst for some vars before output???

            puts $fr_id $output_tmp_line

            unset output_tmp_line
            unset ttmp_line
         }
      }


      if { $page($page_ind,table) == 0 && $page($page_ind,line) == 0 } {
         # the height of head is too big, so, the first page cannot output anything about table.
      } else {

         for { } { $ta_ind <= $page($page_ind,table) } { } {
           # table title
            for { set ti_ind $ex_doc_table_title($ta_ind,start) } { $ti_ind <= $ex_doc_table_title($ta_ind,end) } { incr ti_ind } {
              #<01-19-11 lxy>
               if { $ta_ind > 0 } {
                  if { [info exists ex_doc_table_title($ta_ind,head_line_num)] &&\
                       $ex_doc_table_title($ta_ind,head_line_num) == $ti_ind } {
                     if { ![info exists ex_doc_table_title($ta_ind,output_head_line)] } {
                        set ex_doc_table_title($ta_ind,head_line_org) $ex_doc_file_of_table($ti_ind)
                        set ex_doc_file_of_table($ti_ind) $ex_doc_table_title($ta_ind,head_line)
                        set ex_doc_table_title($ta_ind,output_head_line) 0
                     } else {
                        set ex_doc_file_of_table($ti_ind) $ex_doc_table_title($ta_ind,head_line_org)
                     }
                  }
               }
               puts $fr_id $ex_doc_file_of_table($ti_ind)
            }

            set var_list $ex_doc_table_line($ta_ind,var_list)
            set first_var [lindex $var_list 0]

           #<Apr-13-2016 gsl> Use data of tool names list to build tool numbers list
            set first_var_org $first_var
            if [string match "mom_tool_number" $first_var] { set first_var "mom_tool_name" }

            set list_var "${first_var}_list"

           # Table line
            for { set line_num $frist_line_num } { $line_num <= $tabel($ta_ind,line_num) } { incr line_num } {

              # Initialize vars used in excel template
               set my_index $line_num
               set $first_var [lindex [set $list_var] [expr $line_num - 1]]

              #<Apr-13-2016 gsl> Critical fix to display proper tool numbers in the 1st column of tool numbers list
              #<Aug-10-2016 shuai> Fix PR7218802. Fix the output problem when the operation tool name is NONE.
               if [string match "mom_tool_number" $first_var_org] {
                  set tmp_var  [lindex $var_list 0]
                  if { [info exists mom_tool_name_data([set $first_var],mom_tool_number)] } {
                     set $tmp_var $mom_tool_name_data([set $first_var],mom_tool_number)
                  } else {
                     set $tmp_var "--"
                  }
               }

               for { set var_ind 1 } { $var_ind < [llength $var_list] } { incr var_ind } {
                  set tmp_var  [lindex $var_list $var_ind]

                  if { [info exists SD_data_list([set $first_var])] } {

                    #<Jan-04-2018 gsl> Direct access to the globals here will only get the last version of the vars. NG!
                     set $tmp_var [lindex $SD_data_list([set $first_var]) [expr $var_ind - 1]]

                  } else {
                     set mom_tool_name "--"
                     set $tmp_var "--"
                  }
               }

               for { set li_ind $ex_doc_table_line($ta_ind,start) } { $li_ind <= $ex_doc_table_line($ta_ind,end) } { incr li_ind } {

                  if { [info exists ex_doc_file_of_table($li_ind)] } {

                     set ttmp_line $ex_doc_file_of_table($li_ind)

                     regsub -all {\[} $ttmp_line {\\[}  ttmp_line
                     regsub -all {\]} $ttmp_line {\\]}  ttmp_line

                    #<Jan-04-2018 gsl> All cells will be evaluated here, since EXP: expression is not created any more.
                     if { ![string match "*\{EXP:*" $ttmp_line] &&\
                           [info exists ex_doc_file_of_table($li_ind,exp)] } {

                        set ide [string last "</td>" $ttmp_line]
                        if { $ide > 0 } {
                           set str [string range $ttmp_line 0 $ide]
                           set ids [string last ">" $str]
                           if { $ids > 0 } {
                              set tts [string range $ttmp_line 0 $ids]
                              set tte [string range $ttmp_line $ide end]

                              set exp $ex_doc_file_of_table($li_ind,exp)

                              if { ![catch { subst $exp }] } {
                                 set exp [subst $exp]
                                 if { ![catch { expr $exp }] } { ;# Not to format string
                                    set fmt [DOC_ask_format_of_style $ex_doc_file_of_table($li_ind,sty)]

                                   #<07-19-2019 gsl> It's possible a string containing all digits would pass expr.
                                    if { $fmt != "" } {
                                       set exp [format "$fmt" [expr ${exp}]]
                                    }
                                 }
                              }

                              set ttmp_line ${tts}${exp}${tte}
                           }
                        }
                     }


                    # Substitute variables
                    #<08-22-2014 gsl> Error protect
                     if ![catch { subst $ttmp_line }] {
                        set ttmp_line [subst $ttmp_line]
                     }
                     regsub -all {\\\[} $ttmp_line {[}  ttmp_line
                     regsub -all {\\\]} $ttmp_line {]}  ttmp_line

                    # Eval expression, if any
                     set ids [string first "\{EXP:" $ttmp_line]
                     if { $ids > 0 } {
                        set ide [string last "\}" $ttmp_line]
                        set ise [expr [string wordend $ttmp_line [expr $ids+5]] - 1]
                        set stl [string range $ttmp_line [expr $ids+5] $ise] ;# style
                        set exp [string range $ttmp_line [expr $ise+2] [expr $ide-1]]

                        set fmt [DOC_ask_format_of_style $stl]

                       # Eval exp
                       #<07-19-2019 gsl> Added condition -
                        if { $fmt != "" } {
                           set exp [format "$fmt" [expr ${exp}]]
                        }

                       # Reconstruct line
                        set tts [string range $ttmp_line 0 [expr $ids-1]]
                        set tte [string range $ttmp_line [expr $ide+1] end]
                        set ttmp_line ${tts}${exp}${tte}

                     } else {
                       #++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                       #<02-26-2014 gsl> It might be too late, since vars have been substituted already!!!
                       #<02-26-2014 gsl> Format variables here
                       # DOC_subst_exist_var ttmp_line
                     }


                    # Add word-wrap attribute to (all for now) cell properties (style=')
                     set ids [string first "style=\'" $ttmp_line]
                     if { $ids > 0 } {
                        regsub "style=\'" $ttmp_line "style=\'word-wrap\:break-word\;" ttmp_line
                     }

                    # regsub -all {\?*} $ttmp_line {} ttmp_line

                     puts $fr_id $ttmp_line

                     unset ttmp_line
                  }
               }

               if { $ta_ind == $page($page_ind,table) && $line_num == $page($page_ind,line) } {
                  set is_break 1
                  set frist_line_num [expr $line_num + 1]

                  if { $line_num == $tabel($ta_ind,line_num) } {

                     incr ta_ind
                     set frist_line_num 1
                  }
                  break
               }
            }

            if { $is_break } {
               set is_break 0
               break
            }

            if { [info exists tabel($ta_ind,line_num)] && $line_num == [expr $tabel($ta_ind,line_num) + 1] } {

               incr ta_ind
               set frist_line_num 1
            }
         }
      }


     # Page footer
      foreach item [lsort -integer [array names ex_doc_file_af_table_end]] {
         set ttmp_line $ex_doc_file_af_table_end($item)
         regsub -all {\[} $ttmp_line {\\[}  ttmp_line
         regsub -all {\]} $ttmp_line {\\]}  ttmp_line
         if { [catch { set output_tmp_line "[subst $ttmp_line]"} res] } {
            set output_tmp_line "$ttmp_line"
         }
         regsub -all {\\\\\[} $output_tmp_line {[}  output_tmp_line
         regsub -all {\\\\\]} $output_tmp_line {]}  output_tmp_line
         puts $fr_id $output_tmp_line
         unset output_tmp_line
         unset ttmp_line
      }
   }


   foreach item [lsort -integer [array names ex_doc_file_af_body_end]] {
      puts $fr_id $ex_doc_file_af_body_end($item)
   }

  # close result file
   close $fr_id

   set cur_dir [pwd]

  #<04-02-11> control the structure of the output files.
   global ex_doc_output_file_structure
   global ex_doc_output_dir
   global ex_doc_template_file

   if { !$ex_doc_output_file_structure } {

     #<Apr-11-2016 gsl> Template may not have associated folder
      if [file exists "[file rootname $ex_doc_template_file]_files"] {
         cd "[file rootname $ex_doc_template_file]_files"
         foreach gif [glob -nocomplain -- "*.png"] {
            file copy -force $gif "$mom_output_file_directory"
         }
         foreach gif [glob -nocomplain -- "*.gif"] {
            file copy -force $gif "$mom_output_file_directory"
         }
      }

   } else {

      cd $mom_output_file_directory
      foreach gif [glob -nocomplain -- "*_path.gif"] {
         file copy -force $gif $ex_doc_output_dir
         file delete -force $gif
      }

     #<Apr-11-2016 gsl> Template may not have associated folder
      if [file exists "[file rootname $ex_doc_template_file]_files"] {
         cd "[file rootname $ex_doc_template_file]_files"
         foreach gif [glob -nocomplain -- "*.png"] {
            file copy -force $gif $ex_doc_output_dir
         }
         foreach gif [glob -nocomplain -- "*.gif"] {
            file copy -force $gif $ex_doc_output_dir
         }
      }
   }

  #<01-07-11 lxy> rollback to the original directory, in case of locking the other folder.
   cd $cur_dir

  #<03-25-11 lxy> place all the output files in the same folder
  #<04-02-11> control the structure of the output files.
   if { $ex_doc_output_file_structure == 2 } {
      set __doc_output_file $ex_doc_output_dir
      append __doc_output_file "/${mom_output_file_basename}.${mom_sys_output_file_suffix}"
      file rename -force $output_file $__doc_output_file
   } else {
      set __doc_output_file $output_file
   }


  #<02-10-11 gsl> Only display the results automatically for Windows -
  #
   if { [string match "windows*" $::tcl_platform(platform)] } {
      if { [info exists ::execute_file] && [file exists $::execute_file] } {
         INFO "Displaying \"$__doc_output_file\" with \"$::execute_file\""
         EXEC "\"$::execute_file\" \"$__doc_output_file\"" 0
      }
   }
}


# #=============================================================
# proc MOM_TOOL_BODY { } {
# #=============================================================
#     global mom_tool_name_data
#     global mom_tool_name_list
#     global mom_tool_number_data
#     global mom_tool_number_list
#     global mom_tool_name
#     global mom_tool_number
#     global mom_tool_type


#    #<10-06-2014 gsl> Enhance operation list with tool data
#     if { [llength [info commands DOC__patch_oper_tool_data]] > 0 } {
#        DOC__patch_oper_tool_data
#     }


#    #<05-18-11 lxy> Some mom variables need to be specially handled for Turning tool.
#     global mom_tracking_point_adjust_register
#     global mom_tool_length_adjust_register
#     global mom_tracking_point_cutcom_register
#     global mom_tool_cutcom_register

#     if { [string match "Turning*" $mom_tool_type] || \
#          [string match "Grooving*" $mom_tool_type] || \
#          [string match "Form*" $mom_tool_type] || \
#          [string match "Threading*" $mom_tool_type] } {

#        if { [info exists mom_tracking_point_adjust_register] } {
#           set mom_tool_length_adjust_register $mom_tracking_point_adjust_register
#        }

#        if { [info exists mom_tracking_point_cutcom_register] } {
#           set mom_tool_cutcom_register $mom_tracking_point_cutcom_register
#        }
#     }

#     global tool_var_list_1  ;# arranged by tool name

#     if { [llength $tool_var_list_1] > 0 } {

#        if { [lsearch $mom_tool_name_list $mom_tool_name] < 0 } {
#           lappend mom_tool_name_list $mom_tool_name

#           foreach tool_var $tool_var_list_1 {

#             #<03-27-2013 gsl> Fixed logic
#              if { [string compare "mom_operation_name" $tool_var] == 0 } {

#                    if { ![info exists mom_tool_name_data($mom_tool_name,$tool_var)] } {
#                       set mom_tool_name_data($mom_tool_name,$tool_var) ""
#                    }

#              } elseif { [string compare "mom_toolpath_time"         $tool_var] == 0 ||\
#                         [string compare "mom_toolpath_cutting_time" $tool_var] == 0 } {

#                    if { ![info exists mom_tool_name_data($mom_tool_name,$tool_var)] } {
#                       set mom_tool_name_data($mom_tool_name,$tool_var) 0
#                    }

#              } else {

#                    if { [info exists ::$tool_var] } {
#                       set mom_tool_name_data($mom_tool_name,$tool_var) [set ::$tool_var]

#                      #<04-13-2016 gsl> Don't unset var, next list needs it
#                      if 0 {
#                      #<03-25-11 lxy> unset all the variables
#                       if { [string compare "mom_tool_name" $tool_var] &&\
#                            [string compare "mom_group_name" $tool_var] } {
#                          unset ::$tool_var
#                       }
#                      }
#                    } else {
#                       set mom_tool_name_data($mom_tool_name,$tool_var) "--"
#                    }
#              }

#           } ;# foreach
#        }

#     }

#    #<03-25-11 lxy> unset all the related variables
#     global oper_var_list
#     foreach oper_var $oper_var_list {
#       # global $oper_var
#        if { [info exists ::$oper_var] && [string compare "mom_group_name" $oper_var] } {
#           unset ::$oper_var
#        }
#     }
# }


#=============================================================
proc MOM_TOOL_BODY { } {
#=============================================================
    # Debug mode (1 - enable, 0 - disable)
    set DEBUG_MODE_TOOL 0

    global mom_tool_name_data
    global mom_tool_name_list
    global mom_tool_number_data
    global mom_tool_number_list
    global mom_tool_name
    global mom_tool_number
    global mom_tool_type


   #<10-06-2014 gsl> Enhance operation list with tool data
    if { [llength [info commands DOC__patch_oper_tool_data]] > 0 } {
       DOC__patch_oper_tool_data
    }


   #<05-18-11 lxy> Some mom variables need to be specially handled for Turning tool.
    global mom_tracking_point_adjust_register
    global mom_tool_length_adjust_register
    global mom_tracking_point_cutcom_register
    global mom_tool_cutcom_register

    if { [string match "Turning*" $mom_tool_type] || \
         [string match "Grooving*" $mom_tool_type] || \
         [string match "Form*" $mom_tool_type] || \
         [string match "Threading*" $mom_tool_type] } {

       if { [info exists mom_tracking_point_adjust_register] } {
          set mom_tool_length_adjust_register $mom_tracking_point_adjust_register
       }

       if { [info exists mom_tracking_point_cutcom_register] } {
          set mom_tool_cutcom_register $mom_tracking_point_cutcom_register
       }
    }

    global tool_var_list_1  ;# arranged by tool name

    if { [llength $tool_var_list_1] > 0 } {

       if { [lsearch $mom_tool_name_list $mom_tool_name] < 0 } {
          lappend mom_tool_name_list $mom_tool_name

          foreach tool_var $tool_var_list_1 {

            #<03-27-2013 gsl> Fixed logic
             if { [string compare "mom_operation_name" $tool_var] == 0 } {

                   if { ![info exists mom_tool_name_data($mom_tool_name,$tool_var)] } {
                      set mom_tool_name_data($mom_tool_name,$tool_var) ""
                   }

             } elseif { [string compare "mom_toolpath_time"         $tool_var] == 0 ||\
                        [string compare "mom_toolpath_cutting_time" $tool_var] == 0 } {

                   if { ![info exists mom_tool_name_data($mom_tool_name,$tool_var)] } {
                      set mom_tool_name_data($mom_tool_name,$tool_var) 0
                   }

             } else {

                   if { [info exists ::$tool_var] } {
                      set mom_tool_name_data($mom_tool_name,$tool_var) [set ::$tool_var]

                     #<04-13-2016 gsl> Don't unset var, next list needs it
                     if 0 {
                     #<03-25-11 lxy> unset all the variables
                      if { [string compare "mom_tool_name" $tool_var] &&\
                           [string compare "mom_group_name" $tool_var] } {
                         unset ::$tool_var
                      }
                     }
                   } else {
                      set mom_tool_name_data($mom_tool_name,$tool_var) "--"
                   }
             }

          } ;# foreach

          # ========== ОТЛАДОЧНЫЙ ВЫВОД (по аналогии с рабочей функцией) ==========
          if { $DEBUG_MODE_TOOL } {
              MOM_output_to_listing_device "\n╔══════════════════════════════════════════════════════════════════╗"
              MOM_output_to_listing_device "║ NEW TOOL ADDED: $mom_tool_name"
              MOM_output_to_listing_device "╚══════════════════════════════════════════════════════════════════╝"
              
              MOM_output_to_listing_device "\n▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄"
              MOM_output_to_listing_device "█ TOOL DATA: $mom_tool_name"
              MOM_output_to_listing_device "▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀"
              
              # Выводим все параметры инструмента
              set tool_vars [array names mom_tool_name_data "$mom_tool_name,*"]
              foreach var_key [lsort $tool_vars] {
                  set var_name [lindex [split $var_key ,] 1]
                  set var_value $mom_tool_name_data($var_key)
                  MOM_output_to_listing_device [format "  %-35s: %s" $var_name $var_value]
              }
              
              MOM_output_to_listing_device "\n════════════════════════════════════════════════════════════════════"
              MOM_output_to_listing_device " END OF TOOL DATA "
              MOM_output_to_listing_device "════════════════════════════════════════════════════════════════════\n"
          }
          # ========== КОНЕЦ ОТЛАДОЧНОГО ВЫВОДА ==========

       } else {
          # Инструмент уже существует - можно вывести короткое уведомление (опционально)
          if { $DEBUG_MODE_TOOL } {
              MOM_output_to_listing_device "  ℹ Tool already exists: $mom_tool_name (skipped duplicate)"
          }
       }

    }

   #<03-25-11 lxy> unset all the related variables
    global oper_var_list
    foreach oper_var $oper_var_list {
      # global $oper_var
       if { [info exists ::$oper_var] && [string compare "mom_group_name" $oper_var] } {
          unset ::$oper_var
       }
    }
}


#=============================================================
proc MOM_OPER_HDR { } {
#=============================================================
}



# Кастомная функция фильтрации массива mom_operation_name_data
# ======================================================================================================================================================================================
# ======================================================================================================================================================================================



# proc MOM_OPER_BODY { } {
#     # Глобальные переменные
#     global mom_operation_name_data
#     global mom_operation_name_list
#     global mom_operation_name
#     global ex_doc_need_capture_path_gif
#     global oper_var_list

#     # Добавляем глобальные переменные для фильтрации
#     global seen_combinations
#     if {![info exists seen_combinations]} {
#         set seen_combinations [list]
#     }

#     # Проверяем уникальность комбинации mom_tool_number и mom_oper_program
#     set tool_number ""
#     set oper_program ""

#     if { [info exists ::mom_tool_number] } {
#         set tool_number $::mom_tool_number
#     }

#     if { [info exists ::mom_oper_program] } {
#         set oper_program $::mom_oper_program
#     }

#     set combination "$tool_number,$oper_program"

#     # Если комбинация уже встречалась, пропускаем операцию
#     if { [lsearch -exact $seen_combinations $combination] != -1 } {
#         return
#     }

#     # Если комбинация уникальна, добавляем её в список seen_combinations
#     lappend seen_combinations $combination

#     # Добавляем операцию в список, если её ещё нет
#     if { [lsearch $mom_operation_name_list $mom_operation_name] < 0 } {
#         lappend mom_operation_name_list $mom_operation_name

#         # Проверяем, есть ли в массиве операция с таким же mom_oper_tool
#         set found_previous_operation 0
#         set previous_tool_description ""
#         set previous_tool_extension_length ""

#         foreach existing_operation $mom_operation_name_list {
#             if { [info exists mom_operation_name_data($existing_operation,mom_oper_tool)] && \
#                  [string equal $mom_operation_name_data($existing_operation,mom_oper_tool) $::mom_oper_tool] } {
#                 # Нашли операцию с таким же mom_oper_tool
#                 set found_previous_operation 1
#                 set previous_tool_description $mom_operation_name_data($existing_operation,mom_tool_description)
#                 set previous_tool_extension_length $mom_operation_name_data($existing_operation,mom_tool_extension_length)
#                 break
#             }
#         }

#         # Если нашли операцию с таким же mom_oper_tool, используем её значения
#         if { $found_previous_operation } {
#             set ::mom_tool_description $previous_tool_description
#             set ::mom_tool_extension_length $previous_tool_extension_length
#         }

#         # Сохраняем данные об операции
#         foreach oper_var $oper_var_list {
#             if { [info exists ::$oper_var] } {
#                 set mom_operation_name_data($mom_operation_name,$oper_var) [set ::$oper_var]
#             } else {
#                 set mom_operation_name_data($mom_operation_name,$oper_var) "--"
#             }
#         }

#         # Очищаем переменные, кроме mom_operation_name и mom_group_name
#         foreach oper_var $oper_var_list {
#             if { [string compare "mom_operation_name" $oper_var] && 
#                  [string compare "mom_group_name" $oper_var] } {
#                 if { [info exists ::$oper_var] } {
#                     unset ::$oper_var
#                 }
#             }
#         }

#         # Форматируем mom_oper_program в массиве
#         if { [info exists mom_operation_name_data($mom_operation_name,mom_oper_program)] } {
#             # Удаляем все "-" и "_", заменяем их на пробелы
#             set formatted_program [string map {"_" " "} $mom_operation_name_data($mom_operation_name,mom_oper_program)]
            
#             # Разделяем formatted_program на две части по первому пробелу
#             set first_space_index [string first " " $formatted_program]
#             if { $first_space_index != -1 } {
#                 # Первая часть (до пробела)
#                 set mom_oper_program_short [string range $formatted_program 0 [expr $first_space_index - 1]]
                
#                 # Вторая часть (после пробела)
#                 set mom_oper_description [string range $formatted_program [expr $first_space_index + 1] end]
                
#                 # Приводим mom_oper_description к нижнему регистру
#                 set mom_oper_description [string tolower $mom_oper_description]
#             } else {
#                 # Если пробела нет, оставляем всё в mom_oper_program
#                 set mom_oper_program_short $formatted_program
#                 set mom_oper_description "--"
#             }

#             # Перезаписываем значения в массиве
#             set mom_operation_name_data($mom_operation_name,mom_oper_program) $mom_oper_program_short
#             set mom_operation_name_data($mom_operation_name,mom_oper_description) $mom_oper_description
#         }
#     }

#     # Остальная часть процедуры (сбор данных об инструментах и времени обработки)
#     global tool_var_list_1
#     global mom_tool_name_data
#     global mom_oper_tool
#     global mom_tool_number

#     if { [lsearch $tool_var_list_1 "mom_operation_name"] >= 0 } {
#         if { [info exists mom_oper_tool] } {
#             if { ![info exists mom_tool_name_data($mom_oper_tool,mom_operation_name)] } {
#                 set mom_tool_name_data($mom_oper_tool,mom_operation_name) "$mom_operation_name"
#             } elseif { [lsearch [split $mom_tool_name_data($mom_oper_tool,mom_operation_name) \n] $mom_operation_name] < 0 } {
#                 append mom_tool_name_data($mom_oper_tool,mom_operation_name) "\n$mom_operation_name"
#             }
#         }
#     }

#     global OPER_tool_number_list OPER_tool_name_list
#     if ![info exists OPER_tool_number_list] { set OPER_tool_number_list [list] }
#     if ![info exists OPER_tool_name_list]   { set OPER_tool_name_list   [list] }

#     if { [info exists mom_tool_number] } {
#         if { [lsearch $OPER_tool_number_list $mom_tool_number] < 0 } {
#             lappend OPER_tool_number_list $mom_tool_number
#         }
#     }

#     if { [info exists mom_oper_tool] } {
#         if { [lsearch $OPER_tool_name_list $mom_oper_tool] < 0 } {
#             lappend OPER_tool_name_list $mom_oper_tool
#         }
#     }
# }




proc MOM_OPER_BODY { } {
    # Глобальные переменные
    global mom_operation_name_data
    global mom_operation_name_list
    global mom_operation_name
    global ex_doc_need_capture_path_gif
    global oper_var_list

    # Добавляем глобальные переменные для фильтрации
    global seen_combinations
    if {![info exists seen_combinations]} {
        set seen_combinations [list]
    }
    
    # ========== ПЕРЕМЕННЫЕ ДЛЯ ВЫВОДА ==========
    global header_printed
    global pending_operation_name
    global pending_operation_printed
    
    if { ![info exists header_printed] } {
        set header_printed 0
    }
    if { ![info exists pending_operation_printed] } {
        set pending_operation_printed 0
    }
    # ===========================================

    # Проверяем уникальность комбинации mom_tool_number и mom_oper_program
    set tool_number ""
    set oper_program ""

    if { [info exists ::mom_tool_number] } {
        set tool_number $::mom_tool_number
    }

    if { [info exists ::mom_oper_program] } {
        set oper_program $::mom_oper_program
    }

    set combination "$tool_number,$oper_program"

    # ========== ВЫВОДИМ ПРЕДЫДУЩУЮ ОПЕРАЦИЮ (только один раз) ==========
    if { $header_printed && [info exists pending_operation_name] && $pending_operation_name != "" && !$pending_operation_printed } {
        # Проверяем, что данные операции уже сохранены и имеют данные инструмента
        if { [info exists mom_operation_name_data($pending_operation_name,mom_tool_number)] } {
            MOM_output_to_listing_device "\n▶ Операция: $pending_operation_name"
            MOM_output_to_listing_device "  ───────────────────────────────────────────────────────────"
            
            set oper_vars [array names mom_operation_name_data "$pending_operation_name,*"]
            foreach var_key [lsort $oper_vars] {
                set var_name [lindex [split $var_key ,] 1]
                set var_value $mom_operation_name_data($var_key)
                MOM_output_to_listing_device [format "  %-30s: %s" $var_name $var_value]
            }
            
            # Отмечаем, что операция уже выведена
            set pending_operation_printed 1
        }
    }
    # ===========================================

    # Если комбинация уже встречалась, пропускаем операцию
    if { [lsearch -exact $seen_combinations $combination] != -1 } {
        return
    }

    # Если комбинация уникальна, добавляем её в список seen_combinations
    lappend seen_combinations $combination

    # Добавляем операцию в список, если её ещё нет
    if { [lsearch $mom_operation_name_list $mom_operation_name] < 0 } {
        lappend mom_operation_name_list $mom_operation_name

        # Сохраняем данные об операции
        foreach oper_var $oper_var_list {
            if { [info exists ::$oper_var] } {
                set mom_operation_name_data($mom_operation_name,$oper_var) [set ::$oper_var]
            } else {
                set mom_operation_name_data($mom_operation_name,$oper_var) "--"
            }
        }

        # Очищаем переменные
        foreach oper_var $oper_var_list {
            if { [string compare "mom_operation_name" $oper_var] && 
                 [string compare "mom_group_name" $oper_var] } {
                if { [info exists ::$oper_var] } {
                    unset ::$oper_var
                }
            }
        }

        # Форматируем mom_oper_program
        if { [info exists mom_operation_name_data($mom_operation_name,mom_oper_program)] } {
            set formatted_program [string map {"_" " "} $mom_operation_name_data($mom_operation_name,mom_oper_program)]
            set first_space_index [string first " " $formatted_program]
            if { $first_space_index != -1 } {
                set mom_oper_program_short [string range $formatted_program 0 [expr $first_space_index - 1]]
                set mom_oper_description [string tolower [string range $formatted_program [expr $first_space_index + 1] end]]
            } else {
                set mom_oper_program_short $formatted_program
                set mom_oper_description "--"
            }
            set mom_operation_name_data($mom_operation_name,mom_oper_program) $mom_oper_program_short
            set mom_operation_name_data($mom_operation_name,mom_oper_description) $mom_oper_description
        }
        
        # Выводим заголовок только один раз
        if { !$header_printed } {
            MOM_output_to_listing_device "\n"
            MOM_output_to_listing_device "════════════════════════════════════════════════════════════════════"
            MOM_output_to_listing_device "СПИСОК ОПЕРАЦИЙ И ИХ ПАРАМЕТРЫ"
            MOM_output_to_listing_device "════════════════════════════════════════════════════════════════════"
            set header_printed 1
        }
        
        # Откладываем ТЕКУЩУЮ операцию для вывода в следующий раз
        set pending_operation_name $mom_operation_name
        set pending_operation_printed 0
    }
}









# ======================================================================================================================================================================================
# ======================================================================================================================================================================================
# ======================================================================================================================================================================================
# ======================================================================================================================================================================================


#=============================================================
proc MOM_OPER_FTR { } {
#=============================================================
}


#=============================================================
proc MOM_PROGRAMVIEW_HDR { } {
#=============================================================
}


#=============================================================
proc MOM_PROGRAM_BODY { } {
#=============================================================
}


#=============================================================
proc MOM_MEMBERS_HDR { } {
#=============================================================
}


#=============================================================
proc MOM_TOOL_HDR { } {
#=============================================================
}


#=============================================================
proc MOM_TOOL_FTR { } {
#=============================================================
}


#=============================================================
proc MOM_MEMBERS_FTR { } {
#=============================================================
#<11-apr-2019 gsl>
# Substitute mom vars for the title block -

   if { [info exists ::ex_doc_title_vars] && [llength $::ex_doc_title_vars] } {
      foreach itm [lsort -integer [array names ::ex_doc_file_bf_table_str]] {
         set line $::ex_doc_file_bf_table_str($itm)
         if [string match "*mom_*" $line] {
            DOC_subst_exist_var line
            set ::ex_doc_file_bf_table_str($itm) $line
         }
      }
      unset ::ex_doc_title_vars
   }
}


#=============================================================
proc MOM_PROGRAMVIEW_FTR { } {
#=============================================================
}


#=============================================================
proc MOM_SETUP_FTR { } {
#=============================================================
}




#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# Source in user's tcl file.
#
# - The user's Tcl file should reside in the directory defined by
#   UGII_CAM_SHOP_DOC_CUSTOM_DIR, or in user's HOME directory.
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

 set USER_SHOPDOC_DIR [MOM_ask_env_var UGII_CAM_SHOP_DOC_CUSTOM_DIR]
 if { $USER_SHOPDOC_DIR == "" } {
    set USER_SHOPDOC_DIR [MOM_ask_env_var HOME]
 }

 if [info exists shopdoc_user_tcl] { unset shopdoc_user_tcl }

 if { $USER_SHOPDOC_DIR != "" } {
    set shopdoc_user_tcl [file join $USER_SHOPDOC_DIR shopdoc_user.tcl]
 }

 if { [info exists shopdoc_user_tcl] && [file exists $shopdoc_user_tcl] && [file size $shopdoc_user_tcl] } {
    INFO "Source in user's Tcl file : $shopdoc_user_tcl"
    source "$shopdoc_user_tcl"
 }



