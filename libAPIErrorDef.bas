Attribute VB_Name = "libAPIErrorDef"
'APIErrorDef - APIErrorDef.bas
'   Win32 API Error String Definition Module...
'Public domain, taken from "The Waite Group's Visual Basic Source Library"/SAMS Publishing...
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   03/04/99    None        Ken Clark       Incorporated into FiRRe; Added GetErrorText from MSDN Sample code;
'=================================================================================================================================
Option Explicit

'***************************************************************
'* APIErrorDef.bas
'* Created by Brian Shea
'* Defines Names to API Return Codes
'***************************************************************

'***************************************************************
'* The section below is a listing of API or system error titles
'* and numbers, commented with the text.
'***************************************************************
'* Clip the needed definitions out of this file and into your
'* project as needed to support the API calls that you are using
'***************************************************************

Public Const ERROR_SUCCESS = 0 'The operation completed successfully
Public Const ERROR_INVALID_FUNCTION = 1 'Incorrect function.
Public Const ERROR_FILE_NOT_FOUND = 2 'The system cannot find the file specified.
Public Const ERROR_PATH_NOT_FOUND = 3 'The system cannot find the path specified.
Public Const ERROR_TOO_MANY_OPEN_FILES = 4 'The system cannot open the file.
Public Const ERROR_ACCESS_DENIED = 5 'Access is denied.
Public Const ERROR_INVALID_HANDLE = 6 'The handle is invalid.
Public Const ERROR_ARENA_TRASHED = 7 'The storage control blocks were destroyed.
Public Const ERROR_NOT_ENOUGH_MEMORY = 8 'Not enough storage is available to process this command.
Public Const ERROR_INVALID_BLOCK = 9 'The storage control block address is invalid.
Public Const ERROR_BAD_ENVIRONMENT = 10 'The environment is incorrect.
Public Const ERROR_BAD_FORMAT = 11 'An attempt was made to load a program with an incorrect format.
Public Const ERROR_INVALID_ACCESS = 12 'The access code is invalid.
Public Const ERROR_INVALID_DATA = 13 'The data is invalid.
Public Const ERROR_OUTOFMEMORY = 14 'Not enough storage is available to complete this operation.
Public Const ERROR_INVALID_DRIVE = 15 'The system cannot find the drive specified.
Public Const ERROR_CURRENT_DIRECTORY = 16 'The directory cannot be removed.
Public Const ERROR_NOT_SAME_DEVICE = 17 'The system cannot move the file to a different disk drive.
Public Const ERROR_NO_MORE_FILES = 18 'There are no more files.
Public Const ERROR_WRITE_PROTECT = 19 'The media is write protected.
Public Const ERROR_BAD_UNIT = 20 'The system cannot find the device specified.
Public Const ERROR_NOT_READY = 21 'The device is not ready.
Public Const ERROR_BAD_COMMAND = 22 'The device does not recognize the command.
Public Const ERROR_CRC = 23 'Data error (cyclic redundancy check).
Public Const ERROR_BAD_LENGTH = 24 'The program issued a command but the command length is incorrect.
Public Const ERROR_SEEK = 25 'The drive cannot locate a specific area or track on the disk.
Public Const ERROR_NOT_DOS_DISK = 26 'The specified disk or diskette cannot be accessed.
Public Const ERROR_SECTOR_NOT_FOUND = 27 'The drive cannot find the sector requested.
Public Const ERROR_OUT_OF_PAPER = 28 'The printer is out of paper.
Public Const ERROR_WRITE_FAULT = 29 'The system cannot write to the specified device.
Public Const ERROR_READ_FAULT = 30 'The system cannot read from the specified device.
'********************************************************************
'* It continues like this for a while
'* <CLIP>
'********************************************************************
Public Const ERROR_GEN_FAILURE = 31 'A device attached to the system is not functioning.
Public Const ERROR_SHARING_VIOLATION = 32 'The process cannot access the file because it is being used by another process.
Public Const ERROR_LOCK_VIOLATION = 33 'The process cannot access the file because another process has locked a portion of the file.
Public Const ERROR_WRONG_DISK = 34 'The wrong diskette is in the drive. Insert %2 (Volume Serial Number: %3) into drive %1.
Public Const ERROR_SHARING_BUFFER_EXCEEDED = 36 'Too many files opened for sharing.
Public Const ERROR_HANDLE_EOF = 38 'Reached the end of the file.
Public Const ERROR_HANDLE_DISK_FULL = 39 'The disk Is full.
'* Error Codes 40 - 49 Not Listed
Public Const ERROR_NOT_SUPPORTED = 50 'The network request is not supported.
Public Const ERROR_REM_NOT_LIST = 51 'The remote computer is not available.
Public Const ERROR_DUP_NAME = 52 'A duplicate name exists on the network.
Public Const ERROR_BAD_NETPATH = 53 'The network path was not found.
Public Const ERROR_NETWORK_BUSY = 54 'The network Is busy.
Public Const ERROR_DEV_NOT_EXIST = 55 'The specified network resource or device is no longer available.
Public Const ERROR_TOO_MANY_CMDS = 56 'The network BIOS command limit has been reached.
Public Const ERROR_ADAP_HDW_ERR = 57 'A network adapter hardware error occurred.
Public Const ERROR_BAD_NET_RESP = 58 'The specified server cannot perform the requested operation.
Public Const ERROR_UNEXP_NET_ERR = 59 'An unexpected network error occurred.
Public Const ERROR_BAD_REM_ADAP = 60 'The remote adapter is not compatible.
Public Const ERROR_PRINTQ_FULL = 61 'The printer queue is full.
Public Const ERROR_NO_SPOOL_SPACE = 62 'Space to store the file waiting to be printed is not available on the server.
Public Const ERROR_PRINT_CANCELLED = 63 'Your file waiting to be printed was deleted.
Public Const ERROR_NETNAME_DELETED = 64 'The specified network name is no longer available.
Public Const ERROR_NETWORK_ACCESS_DENIED = 65 'Network access is denied.
Public Const ERROR_BAD_DEV_TYPE = 66 'The network resource type is not correct.
Public Const ERROR_BAD_NET_NAME = 67 'The network name cannot be found.
Public Const ERROR_TOO_MANY_NAMES = 68 'The name limit for the local computer network adapter card was exceeded.
Public Const ERROR_TOO_MANY_SESS = 69 'The network BIOS session limit was exceeded.
Public Const ERROR_SHARING_PAUSED = 70 'The remote server has been paused or is in the process of being started.
Public Const ERROR_REQ_NOT_ACCEP = 71 'No more connections can be made to this remote computer at this time because there are already as many connections as the computer can accept.
Public Const ERROR_REDIR_PAUSED = 72 'The specified printer or disk device has been paused.
'* Error Codes 73 - 79 Not Listed
Public Const ERROR_FILE_EXISTS = 80 'The file exists.
'* Error Code 81 Not Listed
Public Const ERROR_CANNOT_MAKE = 82 'The directory or file cannot be created.
Public Const ERROR_FAIL_I24 = 83 'Fail on INT 24.
Public Const ERROR_OUT_OF_STRUCTURES = 84 'Storage to process this request is not available.
Public Const ERROR_ALREADY_ASSIGNED = 85 'The local device name is already in use.
Public Const ERROR_INVALID_PASSWORD = 86 'The specified network password is not correct.
Public Const ERROR_INVALID_PARAMETER = 87 'The parameter is incorrect.
Public Const ERROR_NET_WRITE_FAULT = 88 'A write fault occurred on the network.
Public Const ERROR_NO_PROC_SLOTS = 89 'The system cannot start another process at this time.
'* Error Codes 90 - 99 Not Listed
Public Const ERROR_TOO_MANY_SEMAPHORES = 100 'Cannot create another system semaphore.
Public Const ERROR_EXCL_SEM_ALREADY_OWNED = 101 'The exclusive semaphore is owned by another process.
Public Const ERROR_SEM_IS_SET = 102 'The semaphore is set and cannot be closed.
Public Const ERROR_TOO_MANY_SEM_REQUESTS = 103 'The semaphore cannot be set again.
Public Const ERROR_INVALID_AT_INTERRUPT_TIME = 104 'Cannot request exclusive semaphores at interrupt time.
Public Const ERROR_SEM_OWNER_DIED = 105 'The previous ownership of this semaphore has ended.
Public Const ERROR_SEM_USER_LIMIT = 106 'Insert the diskette for drive %1.
Public Const ERROR_DISK_CHANGE = 107 'The program stopped because an alternate diskette was not inserted.
Public Const ERROR_DRIVE_LOCKED = 108 'The disk is in use or locked by another process.
Public Const ERROR_BROKEN_PIPE = 109 'The pipe has been ended.
Public Const ERROR_OPEN_FAILED = 110 'The system cannot open the device or file specified.
Public Const ERROR_BUFFER_OVERFLOW = 111 'The file name is too long.
Public Const ERROR_DISK_FULL = 112 'There is not enough space on the disk.
Public Const ERROR_NO_MORE_SEARCH_HANDLES = 113 'No more internal file identifiers available.
Public Const ERROR_INVALID_TARGET_HANDLE = 114 'The target internal file identifier is incorrect.
Public Const ERROR_INVALID_CATEGORY = 117 'The IOCTL call made by the application program is not correct.
Public Const ERROR_INVALID_VERIFY_SWITCH = 118 'The verify-on-write switch parameter value is not correct.
Public Const ERROR_BAD_DRIVER_LEVEL = 119 'The system does not support the command requested.
Public Const ERROR_CALL_NOT_IMPLEMENTED = 120 'This function is not supported on this system.
Public Const ERROR_SEM_TIMEOUT = 121 'The semaphore timeout period has expired.
Public Const ERROR_INSUFFICIENT_BUFFER = 122 'The data area passed to a system call is too small.
Public Const ERROR_INVALID_NAME = 123 'The filename, directory name, or volume label syntax is incorrect.
Public Const ERROR_INVALID_LEVEL = 124 'The system call level is not correct.
Public Const ERROR_NO_VOLUME_LABEL = 125 'The disk has no volume label.
Public Const ERROR_MOD_NOT_FOUND = 126 'The specified module could not be found.
Public Const ERROR_PROC_NOT_FOUND = 127 'The specified procedure could not be found.
Public Const ERROR_WAIT_NO_CHILDREN = 128 'There are no child processes to wait for.
Public Const ERROR_CHILD_NOT_COMPLETE = 129 'The %1 application cannot be run in Win32 mode.
Public Const ERROR_DIRECT_ACCESS_HANDLE = 130 'Attempt to use a file handle to an open disk partition for an operation other than raw disk I/O.
Public Const ERROR_NEGATIVE_SEEK = 131 'An attempt was made to move the file pointer before the beginning of the file.
Public Const ERROR_SEEK_ON_DEVICE = 132 'The file pointer cannot be set on the specified device or file.
Public Const ERROR_IS_JOIN_TARGET = 133 'A JOIN or SUBST command cannot be used for a drive that contains previously joined drives.
Public Const ERROR_IS_JOINED = 134 'An attempt was made to use a JOIN or SUBST command on a drive that has already been joined.
Public Const ERROR_IS_SUBSTED = 135 'An attempt was made to use a JOIN or SUBST command on a drive that has already been substituted.
Public Const ERROR_NOT_JOINED = 136 'The system tried to delete the JOIN of a drive that is not joined.
Public Const ERROR_NOT_SUBSTED = 137 'The system tried to delete the substitution of a drive that is not substituted.
Public Const ERROR_JOIN_TO_JOIN = 138 'The system tried to join a drive to a directory on a joined drive.
Public Const ERROR_SUBST_TO_SUBST = 139 'The system tried to substitute a drive to a directory on a substituted drive.
Public Const ERROR_JOIN_TO_SUBST = 140 'The system tried to join a drive to a directory on a substituted drive.
Public Const ERROR_SUBST_TO_JOIN = 141 'The system tried to SUBST a drive to a directory on a joined drive.
Public Const ERROR_BUSY_DRIVE = 142 'The system cannot perform a JOIN or SUBST at this time.
Public Const ERROR_SAME_DRIVE = 143 'The system cannot join or substitute a drive to or for a directory on the same drive.
Public Const ERROR_DIR_NOT_ROOT = 144 'The directory is not a subdirectory of the root directory.
Public Const ERROR_DIR_NOT_EMPTY = 145 'The directory is not empty.
Public Const ERROR_IS_SUBST_PATH = 146 'The path specified is being used in a substitute.
Public Const ERROR_IS_JOIN_PATH = 147 'Not enough resources are available to process this command.
Public Const ERROR_PATH_BUSY = 148 'The path specified cannot be used at this time.
Public Const ERROR_IS_SUBST_TARGET = 149 'An attempt was made to join or substitute a drive for which a directory on the drive is the target of a previous substitute.
Public Const ERROR_SYSTEM_TRACE = 150 'System trace information was not specified in your CONFIG.SYS file, or tracing is disallowed.
Public Const ERROR_INVALID_EVENT_COUNT = 151 'The number of specified semaphore events for DosMuxSemWait is not correct.
Public Const ERROR_TOO_MANY_MUXWAITERS = 152 'DosMuxSemWait did not execute; too many semaphores are already set.
Public Const ERROR_INVALID_LIST_FORMAT = 153 'The DosMuxSemWait list is not correct.
Public Const ERROR_LABEL_TOO_LONG = 154 'The volume label you entered exceeds the label character limit of the target file system.
Public Const ERROR_TOO_MANY_TCBS = 155 'Cannot create another thread.
Public Const ERROR_SIGNAL_REFUSED = 156 'The recipient process has refused the signal.
Public Const ERROR_DISCARDED = 157 'The segment is already discarded and cannot be locked.
Public Const ERROR_NOT_LOCKED = 158 'The segment is already unlocked.
Public Const ERROR_BAD_THREADID_ADDR = 159 'The address for the thread ID is not correct.
Public Const ERROR_BAD_ARGUMENTS = 160 'The argument string passed to DosExecPgm is not correct.
Public Const ERROR_BAD_PATHNAME = 161 'The specified path is invalid.
Public Const ERROR_SIGNAL_PENDING = 162 'A signal is already pending.
'* Error Code 163 Not Listed
Public Const ERROR_MAX_THRDS_REACHED = 164 'No more threads can be created in the system.
'* Error Codes 165 - 166 Not Listed
Public Const ERROR_LOCK_FAILED = 167 'Unable to lock a region of a file.
'* Error Codes 168 - 169 Not Listed
Public Const ERROR_BUSY = 170 'The requested resource is in use.
'* Error Codes 171 - 172 Not Listed
Public Const ERROR_CANCEL_VIOLATION = 173 'A lock request was not outstanding for the supplied cancel region.
Public Const ERROR_ATOMIC_LOCKS_NOT_SUPPORTED = 174 'The file system does not support atomic changes to the lock type.
'* Error Codes 175 - 179 Not Listed
Public Const ERROR_INVALID_SEGMENT_NUMBER = 180 'The system detected a segment number that was not correct.
'* Error Code 181 Nor Listed
Public Const ERROR_INVALID_ORDINAL = 182 'The operating system cannot run %1.
Public Const ERROR_ALREADY_EXISTS = 183 'Cannot create a file when that file already exists.
'* Error Codes 184 - 185 Not Listed
Public Const ERROR_INVALID_FLAG_NUMBER = 186 'The flag passed is not correct.
Public Const ERROR_SEM_NOT_FOUND = 187 'The specified system semaphore name was not found.
Public Const ERROR_INVALID_STARTING_CODESEG = 188 'The operating system cannot run %1.
Public Const ERROR_INVALID_STACKSEG = 189 'The operating system cannot run %1.
Public Const ERROR_INVALID_MODULETYPE = 190 'The operating system cannot run %1.
Public Const ERROR_INVALID_EXE_SIGNATURE = 191 'Cannot run %1 in Win32 mode.
Public Const ERROR_EXE_MARKED_INVALID = 192 'The operating system cannot run %1.
Public Const ERROR_BAD_EXE_FORMAT = 193 '%1 is not a valid Win32 application.
Public Const ERROR_ITERATED_DATA_EXCEEDS_64k = 194 'The operating system cannot run %1.
Public Const ERROR_INVALID_MINALLOCSIZE = 195 'The operating system cannot run %1.
Public Const ERROR_DYNLINK_FROM_INVALID_RING = 196 'The operating system cannot run this application program.
Public Const ERROR_IOPL_NOT_ENABLED = 197 'The operating system is not presently configured to run this application.
Public Const ERROR_INVALID_SEGDPL = 198 'The operating system cannot run %1.
Public Const ERROR_AUTODATASEG_EXCEEDS_64k = 199 'The operating system cannot run this application program.
Public Const ERROR_RING2SEG_MUST_BE_MOVABLE = 200 'The code segment cannot be greater than or equal to 64K.
Public Const ERROR_RELOC_CHAIN_XEEDS_SEGLIM = 201 'The operating system cannot run %1.
Public Const ERROR_INFLOOP_IN_RELOC_CHAIN = 202 'The operating system cannot run %1.
Public Const ERROR_ENVVAR_NOT_FOUND = 203 'The system could not find the environment option that was entered.
'* Error Code 204 Not Listed
Public Const ERROR_NO_SIGNAL_SENT = 205 'No process in the command subtree has a signal handler.
Public Const ERROR_FILENAME_EXCED_RANGE = 206 'The filename or extension is too long.
Public Const ERROR_RING2_STACK_IN_USE = 207 'The ring 2 stack is in use.
Public Const ERROR_META_EXPANSION_TOO_LONG = 208 'The global filename characters, * or ?, are entered incorrectly or too many global filename characters are specified.
Public Const ERROR_INVALID_SIGNAL_NUMBER = 209 'The signal being posted is not correct.
Public Const ERROR_THREAD_1_INACTIVE = 210 'The signal handler cannot be set.
'* Error Code 211 Not Listed
Public Const ERROR_LOCKED = 212 'The segment is locked and cannot be reallocated.
'* Error Code 213 Not Listed
Public Const ERROR_TOO_MANY_MODULES = 214 'Too many dynamic-link modules are attached to this program or dynamic-link module.
Public Const ERROR_NESTING_NOT_ALLOWED = 215 'Can't nest calls to LoadModule.
Public Const ERROR_EXE_MACHINE_TYPE_MISMATCH = 216 'The image file %1 is valid, but is for a machine type other than the current machine.
'* Error Codes 217 - 229 Not Listed
Public Const ERROR_BAD_PIPE = 230 'The pipe state is invalid.
Public Const ERROR_PIPE_BUSY = 231 'All pipe instances are busy.
Public Const ERROR_NO_DATA = 232 'The pipe is being closed.
Public Const ERROR_PIPE_NOT_CONNECTED = 233 'No process is on the other end of the pipe.
Public Const ERROR_MORE_DATA = 234 'More data is available.
'* Error Codes 235 - 239 Not Listed
Public Const ERROR_VC_DISCONNECTED = 240 'The session was canceled.
'* Error Codes 241 - 253 Not Listed
Public Const ERROR_INVALID_EA_NAME = 254 'The specified extended attribute name was invalid.
Public Const ERROR_EA_LIST_INCONSISTENT = 255 'The extended attributes are inconsistent.
'* Error Codes 256 - 258 Not Listed
Public Const ERROR_NO_MORE_ITEMS = 259 'No more data is available.
'* Error Codes 260 - 265 Not Listed
Public Const ERROR_CANNOT_COPY = 266 'The copy functions cannot be used.
Public Const ERROR_DIRECTORY = 267 'The directory name is invalid.
'* Error Codes 268 - 274 Not Listed
Public Const ERROR_EAS_DIDNT_FIT = 275 'The extended attributes did not fit in the buffer.
Public Const ERROR_EA_FILE_CORRUPT = 276 'The extended attribute file on the mounted file system is corrupt.
Public Const ERROR_EA_TABLE_FULL = 277 'The extended attribute table file is full.
Public Const ERROR_INVALID_EA_HANDLE = 278 'The specified extended attribute handle is invalid.
'* Error Codes 279 - 281 Not Listed
Public Const ERROR_EAS_NOT_SUPPORTED = 282 'The mounted file system does not support extended attributes.
'* Error Codes 283 - 287 Not Listed
Public Const ERROR_NOT_OWNER = 288 'Attempt to release mutex not owned by caller.
'* Error Codes 289 - 297 Not Listed
Public Const ERROR_TOO_MANY_POSTS = 298 'Too many posts were made to a semaphore.
Public Const ERROR_PARTIAL_COPY = 299 'Only part of a ReadProcessMemoty or WriteProcessMemory request was completed.
Public Const ERROR_OPLOCK_NOT_GRANTED = 300 'The oplock request is denied.
Public Const ERROR_INVALID_OPLOCK_PROTOCOL = 301 'An invalid oplock acknowledgment was received by the system.
'* Error Codes 302 - 316 Not Listed
Public Const ERROR_MR_MID_NOT_FOUND = 317 'The system cannot find message text for message number 0x%1 in the message file for %2.
'* Error Codes 318 - 486 Not Listed
Public Const ERROR_INVALID_ADDRESS = 487 'Attempt to access invalid address.
'* Error Codes 488 - 533 Not Listed
Public Const ERROR_ARITHMETIC_OVERFLOW = 534 'Arithmetic result exceeded 32 bits.
Public Const ERROR_PIPE_CONNECTED = 535 'There is a process on other end of the pipe.
Public Const ERROR_PIPE_LISTENING = 536 'Waiting for a process to open the other end of the pipe.
'* Error Codes 537 - 993 Not Listed
Public Const ERROR_EA_ACCESS_DENIED = 994 'Access to the extended attribute was denied.
Public Const ERROR_OPERATION_ABORTED = 995 'The I/O operation has been aborted because of either a thread exit or an application request.
Public Const ERROR_IO_INCOMPLETE = 996 'Overlapped I/O event is not in a signaled state.
Public Const ERROR_IO_PENDING = 997 'Overlapped I/O operation is in progress.
Public Const ERROR_NOACCESS = 998 'Invalid access to memory location.
Public Const ERROR_SWAPERROR = 999 'Error performing inpage operation.
'* Error Code 1000 Not Listed
Public Const ERROR_STACK_OVERFLOW = 1001 'Recursion too deep; the stack overflowed.
Public Const ERROR_INVALID_MESSAGE = 1002 'The window cannot act on the sent message.
Public Const ERROR_CAN_NOT_COMPLETE = 1003 'Cannot complete this function.
Public Const ERROR_INVALID_FLAGS = 1004 'Invalid flags.
Public Const ERROR_UNRECOGNIZED_VOLUME = 1005 'The volume does not contain a recognized file system. Please make sure that all required file system drivers are loaded and that the volume is not corrupted.
Public Const ERROR_FILE_INVALID = 1006 'The volume for a file has been externally altered so that the opened file is no longer valid.
Public Const ERROR_FULLSCREEN_MODE = 1007 'The requested operation cannot be performed in full-screen mode.
Public Const ERROR_NO_TOKEN = 1008 'An attempt was made to reference a token that does not exist.
Public Const ERROR_BADDB = 1009 'The configuration registry database is corrupt.
Public Const ERROR_BADKEY = 1010 'The configuration registry key is invalid.
Public Const ERROR_CANTOPEN = 1011 'The configuration registry key could not be opened.
Public Const ERROR_CANTREAD = 1012 'The configuration registry key could not be read.
Public Const ERROR_CANTWRITE = 1013 'The configuration registry key could not be written.
Public Const ERROR_REGISTRY_RECOVERED = 1014 'One of the files in the registry database had to be recovered by use of a log or alternate copy. The recovery was successful.
Public Const ERROR_REGISTRY_CORRUPT = 1015 'The registry is corrupted. The structure of one of the files that contains registry data is corrupted, or the system's image of the file in memory is corrupted, or the file could not be recovered because the alternate copy or log was absent or corrupted.
Public Const ERROR_REGISTRY_IO_FAILED = 1016 'An I/O operation initiated by the registry failed unrecoverably. The registry could not read in, or write out, or flush, one of the files that contain the system's image of the registry.
Public Const ERROR_NOT_REGISTRY_FILE = 1017 'The system has attempted to load or restore a file into the registry, but the specified file is not in a registry file format.
Public Const ERROR_KEY_DELETED = 1018 'Illegal operation attempted on a registry key that has been marked for deletion.
Public Const ERROR_NO_LOG_SPACE = 1019 'System could not allocate the required space in a registry log.
Public Const ERROR_KEY_HAS_CHILDREN = 1020 'Cannot create a symbolic link in a registry key that already has subkeys or values.
Public Const ERROR_CHILD_MUST_BE_VOLATILE = 1021 'Cannot create a stable subkey under a volatile parent key.
Public Const ERROR_NOTIFY_ENUM_DIR = 1022 'A notify change request is being completed and the information is not being returned in the caller's buffer. The caller now needs to enumerate the files to find the changes.
'* Error Codes 1023 - 1050 Not Listed
Public Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051 'A stop control has been sent to a service that other running services are dependent on.
Public Const ERROR_INVALID_SERVICE_CONTROL = 1052 'The requested control is not valid for this service.
Public Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053 'The service did not respond to the start or control request in a timely fashion.
Public Const ERROR_SERVICE_NO_THREAD = 1054 'A thread could not be created for the service.
Public Const ERROR_SERVICE_DATABASE_LOCKED = 1055 'The service database is locked.
Public Const ERROR_SERVICE_ALREADY_RUNNING = 1056 'An instance of the service is already running.
Public Const ERROR_INVALID_SERVICE_ACCOUNT = 1057 'The account name is invalid or does not exist.
Public Const ERROR_SERVICE_DISABLED = 1058 'The service cannot be started, either because it is disabled or because it has no enabled devices associated with it.
Public Const ERROR_CIRCULAR_DEPENDENCY = 1059 'Circular service dependency was specified.
Public Const ERROR_SERVICE_DOES_NOT_EXIST = 1060 'The specified service does not exist as an installed service.
Public Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL = 1061 'The service cannot accept control messages at this time.
Public Const ERROR_SERVICE_NOT_ACTIVE = 1062 'The service has not been started.
Public Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063 'The service process could not connect to the service controller.
Public Const ERROR_EXCEPTION_IN_SERVICE = 1064 'An exception occurred in the service when handling the control request.
Public Const ERROR_DATABASE_DOES_NOT_EXIST = 1065 'The database specified does not exist.
Public Const ERROR_SERVICE_SPECIFIC_ERROR = 1066 'The service has returned a service-specific error code.
Public Const ERROR_PROCESS_ABORTED = 1067 'The process terminated unexpectedly.
Public Const ERROR_SERVICE_DEPENDENCY_FAIL = 1068 'The dependency service or group failed to start.
Public Const ERROR_SERVICE_LOGON_FAILED = 1069 'The service did not start due to a logon failure.
Public Const ERROR_SERVICE_START_HANG = 1070 'After starting, the service hung in a start-pending state.
Public Const ERROR_INVALID_SERVICE_LOCK = 1071 'The specified service database lock is invalid.
Public Const ERROR_SERVICE_MARKED_FOR_DELETE = 1072 'The specified service has been marked for deletion.
Public Const ERROR_SERVICE_EXISTS = 1073 'The specified service already exists.
Public Const ERROR_ALREADY_RUNNING_LKG = 1074 'The system is currently running with the last-known-good configuration.
Public Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075 'The dependency service does not exist or has been marked for deletion.
Public Const ERROR_BOOT_ALREADY_ACCEPTED = 1076 'The current boot has already been accepted for use as the last-known-good control set.
Public Const ERROR_SERVICE_NEVER_STARTED = 1077 'No attempts to start the service have been made since the last boot.
Public Const ERROR_DUPLICATE_SERVICE_NAME = 1078 'The name is already in use as either a service name or a service display name.
Public Const ERROR_DIFFERENT_SERVICE_ACCOUNT = 1079 'The account specified for this service is different from the account specified for other services running in the same process.
Public Const ERROR_CANNOT_DETECT_DRIVER_FAILURE = 1080 'Failure actions can only be set for Win32 services, not for drivers.
Public Const ERROR_CANNOT_DETECT_PROCESS_ABORT = 1081 'This service runs in the same process as the service control manager. Therefore, the service control manager cannot take action if this service's process terminates unexpectedly.
Public Const ERROR_NO_RECOVERY_PROGRAM = 1082 'No recovery program has been configured for this service.
'* Error Codes 1083 - 1099 Not Listed
Public Const ERROR_END_OF_MEDIA = 1100 'The physical end of the tape has been reached.
Public Const ERROR_FILEMARK_DETECTED = 1101 'A tape access reached a filemark.
Public Const ERROR_BEGINNING_OF_MEDIA = 1102 'The beginning of the tape or a partition was encountered.
Public Const ERROR_SETMARK_DETECTED = 1103 'A tape access reached the end of a set of files.
Public Const ERROR_NO_DATA_DETECTED = 1104 'No more data is on the tape.
Public Const ERROR_PARTITION_FAILURE = 1105 'Tape could not be partitioned.
Public Const ERROR_INVALID_BLOCK_LENGTH = 1106 'When accessing a new tape of a multivolume partition, the current blocksize is incorrect.
Public Const ERROR_DEVICE_NOT_PARTITIONED = 1107 'Tape partition information could not be found when loading a tape.
Public Const ERROR_UNABLE_TO_LOCK_MEDIA = 1108 'Unable to lock the media eject mechanism.
Public Const ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109 'Unable to unload the media.
Public Const ERROR_MEDIA_CHANGED = 1110 'The media in the drive may have changed.
Public Const ERROR_BUS_RESET = 1111 'The I/O bus was reset.
Public Const ERROR_NO_MEDIA_IN_DRIVE = 1112 'No media in drive.
Public Const ERROR_NO_UNICODE_TRANSLATION = 1113 'No mapping for the Unicode character exists in the target multi-byte code page.
Public Const ERROR_DLL_INIT_FAILED = 1114 'A dynamic link library (DLL) initialization routine failed.
Public Const ERROR_SHUTDOWN_IN_PROGRESS = 1115 'A system shutdown is in progress.
Public Const ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116 'Unable to abort the system shutdown because no shutdown was in progress.
Public Const ERROR_IO_DEVICE = 1117 'The request could not be performed because of an I/O device error.
Public Const ERROR_SERIAL_NO_DEVICE = 1118 'No serial device was successfully initialized. The serial driver will unload.
Public Const ERROR_IRQ_BUSY = 1119 'Unable to open a device that was sharing an interrupt request (IRQ) with other devices. At least one other device that uses that IRQ was already opened.
Public Const ERROR_MORE_WRITES = 1120 'A serial I/O operation was completed by another write to the serial port. The IOCTL_SERIAL_XOFF_COUNTER reached zero.)
Public Const ERROR_COUNTER_TIMEOUT = 1121 'A serial I/O operation completed because the timeout period expired. The IOCTL_SERIAL_XOFF_COUNTER did not reach zero.)
Public Const ERROR_FLOPPY_ID_MARK_NOT_FOUND = 1122 'No ID address mark was found on the floppy disk.
Public Const ERROR_FLOPPY_WRONG_CYLINDER = 1123 'Mismatch between the floppy disk sector ID field and the floppy disk controller track address.
Public Const ERROR_FLOPPY_UNKNOWN_ERROR = 1124 'The floppy disk controller reported an error that is not recognized by the floppy disk driver.
Public Const ERROR_FLOPPY_BAD_REGISTERS = 1125 'The floppy disk controller returned inconsistent results in its registers.
Public Const ERROR_DISK_RECALIBRATE_FAILED = 1126 'While accessing the hard disk, a recalibrate operation failed, even after retries.
Public Const ERROR_DISK_OPERATION_FAILED = 1127 'While accessing the hard disk, a disk operation failed even after retries.
Public Const ERROR_DISK_RESET_FAILED = 1128 'While accessing the hard disk, a disk controller reset was needed, but even that failed.
Public Const ERROR_EOM_OVERFLOW = 1129 'Physical end of tape encountered.
Public Const ERROR_NOT_ENOUGH_SERVER_MEMORY = 1130 'Not enough server storage is available to process this command.
Public Const ERROR_POSSIBLE_DEADLOCK = 1131 'A potential deadlock condition has been detected.
Public Const ERROR_MAPPED_ALIGNMENT = 1132 'The base address or the file offset specified does not have the proper alignment.
'* Error Codes 1133 - 1139 Not Listed
Public Const ERROR_SET_POWER_STATE_VETOED = 1140 'An attempt to change the system power state was vetoed by another application or driver.
Public Const ERROR_SET_POWER_STATE_FAILED = 1141 'The system BIOS failed an attempt to change the system power state.
Public Const ERROR_TOO_MANY_LINKS = 1142 'An attempt was made to create more links on a file than the file system supports.
'* Error Codes 1143 - 1149 Not Listed
Public Const ERROR_OLD_WIN_VERSION = 1150 'The specified program requires a newer version of Windows.
Public Const ERROR_APP_WRONG_OS = 1151 'The specified program is not a Windows or MS-DOS program.
Public Const ERROR_SINGLE_INSTANCE_APP = 1152 'Cannot start more than one instance of the specified program.
Public Const ERROR_RMODE_APP = 1153 'The specified program was written for an earlier version of Windows.
Public Const ERROR_INVALID_DLL = 1154 'One of the library files needed to run this application is damaged.
Public Const ERROR_NO_ASSOCIATION = 1155 'No application is associated with the specified file for this operation.
Public Const ERROR_DDE_FAIL = 1156 'An error occurred in sending the command to the application.
Public Const ERROR_DLL_NOT_FOUND = 1157 'One of the library files needed to run this application cannot be found.
Public Const ERROR_NO_MORE_USER_HANDLES = 1158 'The current process has used all of its system allowance of handles for Window Manager objects.
Public Const ERROR_MESSAGE_SYNC_ONLY = 1159 'The message can be used only with synchronous operations.
Public Const ERROR_SOURCE_ELEMENT_EMPTY = 1160 'The indicated source element has no media.
Public Const ERROR_DESTINATION_ELEMENT_FULL = 1161 'The indicated destination element already contains media.
Public Const ERROR_ILLEGAL_ELEMENT_ADDRESS = 1162 'The indicated element does not exist.
Public Const ERROR_MAGAZINE_NOT_PRESENT = 1163 'The indicated element is part of a magazine that is not present.
Public Const ERROR_DEVICE_REINITIALIZATION_NEEDED = 1164 'The indicated device requires reinitialization due to hardware errors.
Public Const ERROR_DEVICE_REQUIRES_CLEANING = 1165 'The device has indicated that cleaning is required before further operations are attempted.
Public Const ERROR_DEVICE_DOOR_OPEN = 1166 'The device has indicated that its door is open.
Public Const ERROR_DEVICE_NOT_CONNECTED = 1167 'The device is not connected.
Public Const ERROR_NOT_FOUND = 1168 'Element not found.
Public Const ERROR_NO_MATCH = 1169 'There was no match for the specified key in the index.
Public Const ERROR_SET_NOT_FOUND = 1170 'The property set specified does not exist on the object.
Public Const ERROR_POINT_NOT_FOUND = 1171 'The point passed to GetMouseMovePoints is not in the buffer.
Public Const ERROR_NO_TRACKING_SERVICE = 1172 'The tracking (workstation) service is not running.
Public Const ERROR_NO_VOLUME_ID = 1173 'The Volume ID could not be found.
'* Error Codes 1174 - 1199 Not Listed
Public Const ERROR_BAD_DEVICE = 1200 'The specified device name is invalid.
Public Const ERROR_CONNECTION_UNAVAIL = 1201 'The device is not currently connected but it is a remembered connection.
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202 'An attempt was made to remember a device that had previously been remembered.
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203 'No network provider accepted the given network path.
Public Const ERROR_BAD_PROVIDER = 1204 'The specified network provider name is invalid.
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205 'Unable to open the network connection profile.
Public Const ERROR_BAD_PROFILE = 1206 'The network connection profile is corrupted.
Public Const ERROR_NOT_CONTAINER = 1207 'Cannot enumerate a noncontainer.
Public Const ERROR_EXTENDED_ERROR = 1208 'An extended error has occurred.
Public Const ERROR_INVALID_GROUPNAME = 1209 'The format of the specified group name is invalid.
Public Const ERROR_INVALID_COMPUTERNAME = 1210 'The format of the specified computer name is invalid.
Public Const ERROR_INVALID_EVENTNAME = 1211 'The format of the specified event name is invalid.
Public Const ERROR_INVALID_DOMAINNAME = 1212 'The format of the specified domain name is invalid.
Public Const ERROR_INVALID_SERVICENAME = 1213 'The format of the specified service name is invalid.
Public Const ERROR_INVALID_NETNAME = 1214 'The format of the specified network name is invalid.
Public Const ERROR_INVALID_SHARENAME = 1215 'The format of the specified share name is invalid.
Public Const ERROR_INVALID_PASSWORDNAME = 1216 'The format of the specified password is invalid.
Public Const ERROR_INVALID_MESSAGENAME = 1217 'The format of the specified message name is invalid.
Public Const ERROR_INVALID_MESSAGEDEST = 1218 'The format of the specified message destination is invalid.
Public Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219 'The credentials supplied conflict with an existing set of credentials.
Public Const ERROR_REMOTE_SESSION_LIMIT_EXCEEDED = 1220 'An attempt was made to establish a session to a network server, but there are already too many sessions established to that server.
Public Const ERROR_DUP_DOMAINNAME = 1221 'The workgroup or domain name is already in use by another computer on the network.
Public Const ERROR_NO_NETWORK = 1222 'The network is not present or not started.
Public Const ERROR_CANCELLED = 1223 'The operation was canceled by the user.
Public Const ERROR_USER_MAPPED_FILE = 1224 'The requested operation cannot be performed on a file with a user-mapped section open.
Public Const ERROR_CONNECTION_REFUSED = 1225 'The remote system refused the network connection.
Public Const ERROR_GRACEFUL_DISCONNECT = 1226 'The network connection was gracefully closed.
Public Const ERROR_ADDRESS_ALREADY_ASSOCIATED = 1227 'The network transport endpoint already has an address associated with it.
Public Const ERROR_ADDRESS_NOT_ASSOCIATED = 1228 'An address has not yet been associated with the network endpoint.
Public Const ERROR_CONNECTION_INVALID = 1229 'An operation was attempted on a nonexistent network connection.
Public Const ERROR_CONNECTION_ACTIVE = 1230 'An invalid operation was attempted on an active network connection.
Public Const ERROR_NETWORK_UNREACHABLE = 1231 'The remote network is not reachable by the transport.
Public Const ERROR_HOST_UNREACHABLE = 1232 'The remote system is not reachable by the transport.
Public Const ERROR_PROTOCOL_UNREACHABLE = 1233 'The remote system does not support the transport protocol.
Public Const ERROR_PORT_UNREACHABLE = 1234 'No service is operating at the destination network endpoint on the remote system.
Public Const ERROR_REQUEST_ABORTED = 1235 'The request was aborted.
Public Const ERROR_CONNECTION_ABORTED = 1236 'The network connection was aborted by the local system.
Public Const ERROR_RETRY = 1237 'The operation could not be completed. A retry should be performed.
Public Const ERROR_CONNECTION_COUNT_LIMIT = 1238 'A connection to the server could not be made because the limit on the number of concurrent connections for this account has been reached.
Public Const ERROR_LOGIN_TIME_RESTRICTION = 1239 'Attempting to log in during an unauthorized time of day for this account.
Public Const ERROR_LOGIN_WKSTA_RESTRICTION = 1240 'The account is not authorized to log in from this station.
Public Const ERROR_INCORRECT_ADDRESS = 1241 'The network address could not be used for the operation requested.
Public Const ERROR_ALREADY_REGISTERED = 1242 'The service is already registered.
Public Const ERROR_SERVICE_NOT_FOUND = 1243 'The specified service does not exist.
Public Const ERROR_NOT_AUTHENTICATED = 1244 'The operation being requested was not performed because the user has not been authenticated.
Public Const ERROR_NOT_LOGGED_ON = 1245 'The operation being requested was not performed because the user has not logged on to the network. The specified service does not exist.
Public Const ERROR_CONTINUE = 1246 'Continue with work in progress.
Public Const ERROR_ALREADY_INITIALIZED = 1247 'An attempt was made to perform an initialization operation when initialization has already been completed.
Public Const ERROR_NO_MORE_DEVICES = 1248 'No more local devices.
Public Const ERROR_NO_SUCH_SITE = 1249 'The specified site does not exist.
Public Const ERROR_DOMAIN_CONTROLLER_EXISTS = 1250 'A domain controller with the specified name already exists.
Public Const ERROR_DS_NOT_INSTALLED = 1251 'An error occurred while installing the Windows NT directory service. Please view the event log for more information.
'* Error Codes 1252 - 1299 Not Listed
Public Const ERROR_NOT_ALL_ASSIGNED = 1300 'Not all privileges referenced are assigned to the caller.
Public Const ERROR_SOME_NOT_MAPPED = 1301 'Some mapping between account names and security IDs was not done.
Public Const ERROR_NO_QUOTAS_FOR_ACCOUNT = 1302 'No system quota limits are specifically set for this account.
Public Const ERROR_LOCAL_USER_SESSION_KEY = 1303 'No encryption key is available. A well-known encryption key was returned.
Public Const ERROR_NULL_LM_PASSWORD = 1304 'The Windows NT password is too complex to be converted to a LAN Manager password. The LAN Manager password returned is a NULL string.
Public Const ERROR_UNKNOWN_REVISION = 1305 'The revision level is unknown.
Public Const ERROR_REVISION_MISMATCH = 1306 'Indicates two revision levels are incompatible.
Public Const ERROR_INVALID_OWNER = 1307 'This security ID may not be assigned as the owner of this object.
Public Const ERROR_INVALID_PRIMARY_GROUP = 1308 'This security ID may not be assigned as the primary group of an object.
Public Const ERROR_NO_IMPERSONATION_TOKEN = 1309 'An attempt has been made to operate on an impersonation token by a thread that is not currently impersonating a client.
Public Const ERROR_CANT_DISABLE_MANDATORY = 1310 'The group may not be disabled.
Public Const ERROR_NO_LOGON_SERVERS = 1311 'There are currently no logon servers available to service the logon request.
Public Const ERROR_NO_SUCH_LOGON_SESSION = 1312 'A specified logon session does not exist. It may already have been terminated.
Public Const ERROR_NO_SUCH_PRIVILEGE = 1313 'A specified privilege does not exist.
Public Const ERROR_PRIVILEGE_NOT_HELD = 1314 'A required privilege is not held by the client.
Public Const ERROR_INVALID_ACCOUNT_NAME = 1315 'The name provided is not a properly formed account name.
Public Const ERROR_USER_EXISTS = 1316 'The specified user already exists.
Public Const ERROR_NO_SUCH_USER = 1317 'The specified user does not exist.
Public Const ERROR_GROUP_EXISTS = 1318 'The specified group already exists.
Public Const ERROR_NO_SUCH_GROUP = 1319 'The specified group does not exist.
Public Const ERROR_MEMBER_IN_GROUP = 1320 'Either the specified user account is already a member of the specified group, or the specified group cannot be deleted because it contains a member.
Public Const ERROR_MEMBER_NOT_IN_GROUP = 1321 'The specified user account is not a member of the specified group account.
Public Const ERROR_LAST_ADMIN = 1322 'The last remaining administration account cannot be disabled or deleted.
Public Const ERROR_WRONG_PASSWORD = 1323 'Unable to update the password. The value provided as the current password is incorrect.
Public Const ERROR_ILL_FORMED_PASSWORD = 1324 'Unable to update the password. The value provided for the new password contains values that are not allowed in passwords.
Public Const ERROR_PASSWORD_RESTRICTION = 1325 'Unable to update the password because a password update rule has been violated.
Public Const ERROR_LOGON_FAILURE = 1326 'Logon failure: unknown user name or bad password.
Public Const ERROR_ACCOUNT_RESTRICTION = 1327 'Logon failure: user account restriction.
Public Const ERROR_INVALID_LOGON_HOURS = 1328 'Logon failure: account logon time restriction violation.
Public Const ERROR_INVALID_WORKSTATION = 1329 'Logon failure: user not allowed to log on to this computer.
Public Const ERROR_PASSWORD_EXPIRED = 1330 'Logon failure: the specified account password has expired.
Public Const ERROR_ACCOUNT_DISABLED = 1331 'Logon failure: account currently disabled.
Public Const ERROR_NONE_MAPPED = 1332 'No mapping between account names and security IDs was done.
Public Const ERROR_TOO_MANY_LUIDS_REQUESTED = 1333 'Too many local user identifiers (LUIDs) were requested at one time.
Public Const ERROR_LUIDS_EXHAUSTED = 1334 'No more local user identifiers (LUIDs) are available.
Public Const ERROR_INVALID_SUB_AUTHORITY = 1335 'The subauthority part of a security ID is invalid for this particular use.
Public Const ERROR_INVALID_ACL = 1336 'The access control list (ACL) structure is invalid.
Public Const ERROR_INVALID_SID = 1337 'The security ID structure is invalid.
Public Const ERROR_INVALID_SECURITY_DESCR = 1338 'The security descriptor structure is invalid.
Public Const ERROR_BAD_INHERITANCE_ACL = 1340 'The inherited access control list (ACL) or access control entry (ACE) could not be built.
Public Const ERROR_SERVER_DISABLED = 1341 'The server is currently disabled.
Public Const ERROR_SERVER_NOT_DISABLED = 1342 'The server is currently enabled.
Public Const ERROR_INVALID_ID_AUTHORITY = 1343 'The value provided was an invalid value for an identifier authority.
Public Const ERROR_ALLOTTED_SPACE_EXCEEDED = 1344 'No more memory is available for security information updates.
Public Const ERROR_INVALID_GROUP_ATTRIBUTES = 1345 'The specified attributes are invalid, or incompatible with the attributes for the group as a whole.
Public Const ERROR_BAD_IMPERSONATION_LEVEL = 1346 'Either a required impersonation level was not provided, or the provided impersonation level is invalid.
Public Const ERROR_CANT_OPEN_ANONYMOUS = 1347 'Cannot open an anonymous level security token.
Public Const ERROR_BAD_VALIDATION_CLASS = 1348 'The validation information class requested was invalid.
Public Const ERROR_BAD_TOKEN_TYPE = 1349 'The type of the token is inappropriate for its attempted use.
Public Const ERROR_NO_SECURITY_ON_OBJECT = 1350 'Unable to perform a security operation on an object that has no associated security.
Public Const ERROR_CANT_ACCESS_DOMAIN_INFO = 1351 'Indicates a Windows NT Server could not be contacted or that objects within the domain are protected such that necessary information could not be retrieved.
Public Const ERROR_INVALID_SERVER_STATE = 1352 'The security account manager (SAM) or local security authority (LSA) server was in the wrong state to perform the security operation.
Public Const ERROR_INVALID_DOMAIN_STATE = 1353 'The domain was in the wrong state to perform the security operation.
Public Const ERROR_INVALID_DOMAIN_ROLE = 1354 'This operation is only allowed for the Primary Domain Controller of the domain.
Public Const ERROR_NO_SUCH_DOMAIN = 1355 'The specified domain did not exist.
Public Const ERROR_DOMAIN_EXISTS = 1356 'The specified domain already exists.
Public Const ERROR_DOMAIN_LIMIT_EXCEEDED = 1357 'An attempt was made to exceed the limit on the number of domains per server.
Public Const ERROR_INTERNAL_DB_CORRUPTION = 1358 'Unable to complete the requested operation because of either a catastrophic media failure or a data structure corruption on the disk.
Public Const ERROR_INTERNAL_ERROR = 1359 'The security account database contains an internal inconsistency.
Public Const ERROR_GENERIC_NOT_MAPPED = 1360 'Generic access types were contained in an access mask which should already be mapped to nongeneric types.
Public Const ERROR_BAD_DESCRIPTOR_FORMAT = 1361 'A security descriptor is not in the right format (absolute or self-relative).
Public Const ERROR_NOT_LOGON_PROCESS = 1362 'The requested action is restricted for use by logon processes only. The calling process has not registered as a logon process.
Public Const ERROR_LOGON_SESSION_EXISTS = 1363 'Cannot start a new logon session with an ID that is already in use.
Public Const ERROR_NO_SUCH_PACKAGE = 1364 'A specified authentication package is unknown.
Public Const ERROR_BAD_LOGON_SESSION_STATE = 1365 'The logon session is not in a state that is consistent with the requested operation.
Public Const ERROR_LOGON_SESSION_COLLISION = 1366 'The logon session ID is already in use.
Public Const ERROR_INVALID_LOGON_TYPE = 1367 'A logon request contained an invalid logon type value.
Public Const ERROR_CANNOT_IMPERSONATE = 1368 'Unable to impersonate using a named pipe until data has been read from that pipe.
Public Const ERROR_RXACT_INVALID_STATE = 1369 'The transaction state of a registry subtree is incompatible with the requested operation.
Public Const ERROR_RXACT_COMMIT_FAILURE = 1370 'An internal security database corruption has been encountered.
Public Const ERROR_SPECIAL_ACCOUNT = 1371 'Cannot perform this operation on built-in accounts.
Public Const ERROR_SPECIAL_GROUP = 1372 'Cannot perform this operation on this built-in special group.
Public Const ERROR_SPECIAL_USER = 1373 'Cannot perform this operation on this built-in special user.
Public Const ERROR_MEMBERS_PRIMARY_GROUP = 1374 'The user cannot be removed from a group because the group is currently the user's primary group.
Public Const ERROR_TOKEN_ALREADY_IN_USE = 1375 'The token is already in use as a primary token.
Public Const ERROR_NO_SUCH_ALIAS = 1376 'The specified local group does not exist.
Public Const ERROR_MEMBER_NOT_IN_ALIAS = 1377 'The specified account name is not a member of the local group.
Public Const ERROR_MEMBER_IN_ALIAS = 1378 'The specified account name is already a member of the local group.
Public Const ERROR_ALIAS_EXISTS = 1379 'The specified local group already exists.
Public Const ERROR_LOGON_NOT_GRANTED = 1380 'Logon failure: the user has not been granted the requested logon type at this computer.
Public Const ERROR_TOO_MANY_SECRETS = 1381 'The maximum number of secrets that may be stored in a single system has been exceeded.
Public Const ERROR_SECRET_TOO_LONG = 1382 'The length of a secret exceeds the maximum length allowed.
Public Const ERROR_INTERNAL_DB_ERROR = 1383 'The local security authority database contains an internal inconsistency.
Public Const ERROR_TOO_MANY_CONTEXT_IDS = 1384 'During a logon attempt, the user's security context accumulated too many security IDs.
Public Const ERROR_LOGON_TYPE_NOT_GRANTED = 1385 'Logon failure: the user has not been granted the requested logon type at this computer.
Public Const ERROR_NT_CROSS_ENCRYPTION_REQUIRED = 1386 'A cross-encrypted password is necessary to change a user password.
Public Const ERROR_NO_SUCH_MEMBER = 1387 'A new member could not be added to a local group because the member does not exist.
Public Const ERROR_INVALID_MEMBER = 1388 'A new member could not be added to a local group because the member has the wrong account type.
Public Const ERROR_TOO_MANY_SIDS = 1389 'Too many security IDs have been specified.
Public Const ERROR_LM_CROSS_ENCRYPTION_REQUIRED = 1390 'A cross-encrypted password is necessary to change this user password.
Public Const ERROR_NO_INHERITANCE = 1391 'Indicates an ACL contains no inheritable components.
Public Const ERROR_FILE_CORRUPT = 1392 'The file or directory is corrupted and unreadable.
Public Const ERROR_DISK_CORRUPT = 1393 'The disk structure is corrupted and unreadable.
Public Const ERROR_NO_USER_SESSION_KEY = 1394 'There is no user session key for the specified logon session.
Public Const ERROR_LICENSE_QUOTA_EXCEEDED = 1395 'The service being accessed is licensed for a particular number of connections. No more connections can be made to the service at this time because there are already as many connections as the service can accept.
'* Error Codes 1396 - 1399 Not Listed
Public Const ERROR_INVALID_WINDOW_HANDLE = 1400 'Invalid window handle.
Public Const ERROR_INVALID_MENU_HANDLE = 1401 'Invalid menu handle.
Public Const ERROR_INVALID_CURSOR_HANDLE = 1402 'Invalid cursor handle.
Public Const ERROR_INVALID_ACCEL_HANDLE = 1403 'Invalid accelerator table handle.
Public Const ERROR_INVALID_HOOK_HANDLE = 1404 'Invalid hook handle.
Public Const ERROR_INVALID_DWP_HANDLE = 1405 'Invalid handle to a multiple-window position structure.
Public Const ERROR_TLW_WITH_WSCHILD = 1406 'Cannot create a top-level child window.
Public Const ERROR_CANNOT_FIND_WND_CLASS = 1407 'Cannot find window class.
Public Const ERROR_WINDOW_OF_OTHER_THREAD = 1408 'Invalid window; it belongs to other thread.
Public Const ERROR_HOTKEY_ALREADY_REGISTERED = 1409 'Hot key is already registered.
Public Const ERROR_CLASS_ALREADY_EXISTS = 1410 'Class already exists.
Public Const ERROR_CLASS_DOES_NOT_EXIST = 1411 'Class does not exist.
Public Const ERROR_CLASS_HAS_WINDOWS = 1412 'Class still has open windows.
Public Const ERROR_INVALID_INDEX = 1413 'Invalid index.
Public Const ERROR_INVALID_ICON_HANDLE = 1414 'Invalid icon handle.
Public Const ERROR_PRIVATE_DIALOG_INDEX = 1415 'Using private DIALOG window words.
Public Const ERROR_LISTBOX_ID_NOT_FOUND = 1416 'The list box identifier was not found.
Public Const ERROR_NO_WILDCARD_CHARACTERS = 1417 'No wildcards were found.
Public Const ERROR_CLIPBOARD_NOT_OPEN = 1418 'Thread does not have a clipboard open.
Public Const ERROR_HOTKEY_NOT_REGISTERED = 1419 'Hot key is not registered.
Public Const ERROR_WINDOW_NOT_DIALOG = 1420 'The window is not a valid dialog window.
Public Const ERROR_CONTROL_ID_NOT_FOUND = 1421 'Control ID not found.
Public Const ERROR_INVALID_COMBOBOX_MESSAGE = 1422 'Invalid message for a combo box because it does not have an edit control.
Public Const ERROR_WINDOW_NOT_COMBOBOX = 1423 'The window is not a combo box.
Public Const ERROR_INVALID_EDIT_HEIGHT = 1424 'Height must be less than 256.
Public Const ERROR_DC_NOT_FOUND = 1425 'Invalid device context (DC) handle.
Public Const ERROR_INVALID_HOOK_FILTER = 1426 'Invalid hook procedure type.
Public Const ERROR_INVALID_FILTER_PROC = 1427 'Invalid hook procedure.
Public Const ERROR_HOOK_NEEDS_HMOD = 1428 'Cannot set nonlocal hook without a module handle.
Public Const ERROR_GLOBAL_ONLY_HOOK = 1429 'This hook procedure can only be set globally.
Public Const ERROR_JOURNAL_HOOK_SET = 1430 'The journal hook procedure is already installed.
Public Const ERROR_HOOK_NOT_INSTALLED = 1431 'The hook procedure is not installed.
Public Const ERROR_INVALID_LB_MESSAGE = 1432 'Invalid message for single-selection list box.
Public Const ERROR_SETCOUNT_ON_BAD_LB = 1433 'LB_SETCOUNT sent to non-lazy list box.
Public Const ERROR_LB_WITHOUT_TABSTOPS = 1434 'This list box does not support tab stops.
Public Const ERROR_DESTROY_OBJECT_OF_OTHER_THREAD = 1435 'Cannot destroy object created by another thread.
Public Const ERROR_CHILD_WINDOW_MENU = 1436 'Child windows cannot have menus.
Public Const ERROR_NO_SYSTEM_MENU = 1437 'The window does not have a system menu.
Public Const ERROR_INVALID_MSGBOX_STYLE = 1438 'Invalid message box style.
Public Const ERROR_INVALID_SPI_VALUE = 1439 'Invalid system-wide (SPI_*) parameter.
Public Const ERROR_SCREEN_ALREADY_LOCKED = 1440 'Screen already locked.
Public Const ERROR_HWNDS_HAVE_DIFF_PARENT = 1441 'All handles to windows in a multiple-window position structure must have the same parent.
Public Const ERROR_NOT_CHILD_WINDOW = 1442 'The window is not a child window.
Public Const ERROR_INVALID_GW_COMMAND = 1443 'Invalid GW_ * Command.
Public Const ERROR_INVALID_THREAD_ID = 1444 'Invalid thread identifier.
Public Const ERROR_NON_MDICHILD_WINDOW = 1445 'Cannot process a message from a window that is not a multiple document interface (MDI) window.
Public Const ERROR_POPUP_ALREADY_ACTIVE = 1446 'Popup menu already active.
Public Const ERROR_NO_SCROLLBARS = 1447 'The window does not have scroll bars.
Public Const ERROR_INVALID_SCROLLBAR_RANGE = 1448 'Scroll bar range cannot be greater than 0x7FFF.
Public Const ERROR_INVALID_SHOWWIN_COMMAND = 1449 'Cannot show or remove the window in the way specified.
Public Const ERROR_NO_SYSTEM_RESOURCES = 1450 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_NONPAGED_SYSTEM_RESOURCES = 1451 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_PAGED_SYSTEM_RESOURCES = 1452 'Insufficient system resources exist to complete the requested service.
Public Const ERROR_WORKING_SET_QUOTA = 1453 'Insufficient quota to complete the requested service.
Public Const ERROR_PAGEFILE_QUOTA = 1454 'Insufficient quota to complete the requested service.
Public Const ERROR_COMMITMENT_LIMIT = 1455 'The paging file is too small for this operation to complete.
Public Const ERROR_MENU_ITEM_NOT_FOUND = 1456 'A menu item was not found.
Public Const ERROR_INVALID_KEYBOARD_HANDLE = 1457 'Invalid keyboard layout handle.
Public Const ERROR_HOOK_TYPE_NOT_ALLOWED = 1458 'Hook type not allowed.
Public Const ERROR_REQUIRES_INTERACTIVE_WINDOWSTATION = 1459 'This operation requires an interactive window station.
Public Const ERROR_TIMEOUT = 1460 'This operation returned because the timeout period expired.
Public Const ERROR_INVALID_MONITOR_HANDLE = 1461 'Invalid monitor handle.
'* Error Codes 1462 - 1499 Not Listed
Public Const ERROR_EVENTLOG_FILE_CORRUPT = 1500 'The event log file is corrupted.
Public Const ERROR_EVENTLOG_CANT_START = 1501 'No event log file could be opened, so the event logging service did not start.
Public Const ERROR_LOG_FILE_FULL = 1502 'The event log file is full.
Public Const ERROR_EVENTLOG_FILE_CHANGED = 1503 'The event log file has changed between read operations.
'* Error Codes 1504 - 1600 Not Listed
Public Const ERROR_INSTALL_SERVICE = 1601 'Failure accessing install service.
Public Const ERROR_INSTALL_USEREXIT = 1602 'The user canceled the installation.
Public Const ERROR_INSTALL_FAILURE = 1603 'Fatal error during installation.
Public Const ERROR_INSTALL_SUSPEND = 1604 'Installation suspended, incomplete.
Public Const ERROR_UNKNOWN_PRODUCT = 1605 'Product code not registered.
Public Const ERROR_UNKNOWN_FEATURE = 1606 'Feature ID not registered.
Public Const ERROR_UNKNOWN_COMPONENT = 1607 'Component ID not registered.
Public Const ERROR_UNKNOWN_PROPERTY = 1608 'Unknown property.
Public Const ERROR_INVALID_HANDLE_STATE = 1609 'Handle is in an invalid state.
Public Const ERROR_BAD_CONFIGURATION = 1610 'Configuration data corrupt.
Public Const ERROR_INDEX_ABSENT = 1611 'Language not available.
Public Const ERROR_INSTALL_SOURCE_ABSENT = 1612 'Install source unavailable.
Public Const ERROR_BAD_DATABASE_VERSION = 1613 'Database version unsupported.
Public Const ERROR_PRODUCT_UNINSTALLED = 1614 'Product is uninstalled.
Public Const ERROR_BAD_QUERY_SYNTAX = 1615 'SQL query syntax invalid or unsupported.
Public Const ERROR_INVALID_FIELD = 1616 'Record field does not exist.
'* Error Codes 1617 - 1699 Not Listed
Public Const RPC_S_INVALID_STRING_BINDING = 1700 'The string binding is invalid.
Public Const RPC_S_WRONG_KIND_OF_BINDING = 1701 'The binding handle is not the correct type.
Public Const RPC_S_INVALID_BINDING = 1702 'The binding handle is invalid.
Public Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703 'The RPC protocol sequence is not supported.
Public Const RPC_S_INVALID_RPC_PROTSEQ = 1704 'The RPC protocol sequence is invalid.
Public Const RPC_S_INVALID_STRING_UUID = 1705 'The string universal unique identifier (UUID) is invalid.
Public Const RPC_S_INVALID_ENDPOINT_FORMAT = 1706 'The endpoint format is invalid.
Public Const RPC_S_INVALID_NET_ADDR = 1707 'The network address is invalid.
Public Const RPC_S_NO_ENDPOINT_FOUND = 1708 'No endpoint was found.
Public Const RPC_S_INVALID_TIMEOUT = 1709 'The timeout value is invalid.
Public Const RPC_S_OBJECT_NOT_FOUND = 1710 'The object universal unique identifier (UUID) was not found.
Public Const RPC_S_ALREADY_REGISTERED = 1711 'The object universal unique identifier (UUID) has already been registered.
Public Const RPC_S_TYPE_ALREADY_REGISTERED = 1712 'The type universal unique identifier (UUID) has already been registered.
Public Const RPC_S_ALREADY_LISTENING = 1713 'The RPC server is already listening.
Public Const RPC_S_NO_PROTSEQS_REGISTERED = 1714 'No protocol sequences have been registered.
Public Const RPC_S_NOT_LISTENING = 1715 'The RPC server is not listening.
Public Const RPC_S_UNKNOWN_MGR_TYPE = 1716 'The manager type is unknown.
Public Const RPC_S_UNKNOWN_IF = 1717 'The interface is unknown.
Public Const RPC_S_NO_BINDINGS = 1718 'There are no bindings.
Public Const RPC_S_NO_PROTSEQS = 1719 'There are no protocol sequences.
Public Const RPC_S_CANT_CREATE_ENDPOINT = 1720 'The endpoint cannot be created.
Public Const RPC_S_OUT_OF_RESOURCES = 1721 'Not enough resources are available to complete this operation.
Public Const RPC_S_SERVER_UNAVAILABLE = 1722 'The RPC server is unavailable.
Public Const RPC_S_SERVER_TOO_BUSY = 1723 'The RPC server is too busy to complete this operation.
Public Const RPC_S_INVALID_NETWORK_OPTIONS = 1724 'The network options are invalid.
Public Const RPC_S_NO_CALL_ACTIVE = 1725 'There are no remote procedure calls active on this thread.
Public Const RPC_S_CALL_FAILED = 1726 'The remote procedure call failed.
Public Const RPC_S_CALL_FAILED_DNE = 1727 'The remote procedure call failed and did not execute.
Public Const RPC_S_PROTOCOL_ERROR = 1728 'A remote procedure call (RPC) protocol error occurred.
'* Error Code 1729 Not Listed
Public Const RPC_S_UNSUPPORTED_TRANS_SYN = 1730 'The transfer syntax is not supported by the RPC server.
'* Error Code 1731 Not Listed
Public Const RPC_S_UNSUPPORTED_TYPE = 1732 'The universal unique identifier (UUID) type is not supported.
Public Const RPC_S_INVALID_TAG = 1733 'The tag is invalid.
Public Const RPC_S_INVALID_BOUND = 1734 'The array bounds are invalid.
Public Const RPC_S_NO_ENTRY_NAME = 1735 'The binding does not contain an entry name.
Public Const RPC_S_INVALID_NAME_SYNTAX = 1736 'The name syntax is invalid.
Public Const RPC_S_UNSUPPORTED_NAME_SYNTAX = 1737 'The name syntax is not supported.
'* Error Code 1738 Not Listed
Public Const RPC_S_UUID_NO_ADDRESS = 1739 'No network address is available to use to construct a universal unique identifier (UUID).
Public Const RPC_S_DUPLICATE_ENDPOINT = 1740 'The endpoint is a duplicate.
Public Const RPC_S_UNKNOWN_AUTHN_TYPE = 1741 'The authentication type is unknown.
Public Const RPC_S_MAX_CALLS_TOO_SMALL = 1742 'The maximum number of calls is too small.
Public Const RPC_S_STRING_TOO_LONG = 1743 'The string is too long.
Public Const RPC_S_PROTSEQ_NOT_FOUND = 1744 'The RPC protocol sequence was not found.
Public Const RPC_S_PROCNUM_OUT_OF_RANGE = 1745 'The procedure number is out of range.
Public Const RPC_S_BINDING_HAS_NO_AUTH = 1746 'The binding does not contain any authentication information.
Public Const RPC_S_UNKNOWN_AUTHN_SERVICE = 1747 'The authentication service is unknown.
Public Const RPC_S_UNKNOWN_AUTHN_LEVEL = 1748 'The authentication level is unknown.
Public Const RPC_S_INVALID_AUTH_IDENTITY = 1749 'The security context is invalid.
Public Const RPC_S_UNKNOWN_AUTHZ_SERVICE = 1750 'The authorization service is unknown.
Public Const EPT_S_INVALID_ENTRY = 1751 'The entry is invalid.
Public Const EPT_S_CANT_PERFORM_OP = 1752 'The server endpoint cannot perform the operation.
Public Const EPT_S_NOT_REGISTERED = 1753 'There are no more endpoints available from the endpoint mapper.
Public Const RPC_S_NOTHING_TO_EXPORT = 1754 'No interfaces have been exported.
Public Const RPC_S_INCOMPLETE_NAME = 1755 'The entry name is incomplete.
Public Const RPC_S_INVALID_VERS_OPTION = 1756 'The version option is invalid.
Public Const RPC_S_NO_MORE_MEMBERS = 1757 'There are no more members.
Public Const RPC_S_NOT_ALL_OBJS_UNEXPORTED = 1758 'There is nothing to unexport.
Public Const RPC_S_INTERFACE_NOT_FOUND = 1759 'The interface was not found.
Public Const RPC_S_ENTRY_ALREADY_EXISTS = 1760 'The entry already exists.
Public Const RPC_S_ENTRY_NOT_FOUND = 1761 'The entry is not found.
Public Const RPC_S_NAME_SERVICE_UNAVAILABLE = 1762 'The name service is unavailable.
Public Const RPC_S_INVALID_NAF_ID = 1763 'The network address family is invalid.
Public Const RPC_S_CANNOT_SUPPORT = 1764 'The requested operation is not supported.
Public Const RPC_S_NO_CONTEXT_AVAILABLE = 1765 'No security context is available to allow impersonation.
Public Const RPC_S_INTERNAL_ERROR = 1766 'An internal error occurred in a remote procedure call (RPC).
Public Const RPC_S_ZERO_DIVIDE = 1767 'The RPC server attempted an integer division by zero.
Public Const RPC_S_ADDRESS_ERROR = 1768 'An addressing error occurred in the RPC server.
Public Const RPC_S_FP_DIV_ZERO = 1769 'A floating-point operation at the RPC server caused a division by zero.
Public Const RPC_S_FP_UNDERFLOW = 1770 'A floating-point underflow occurred at the RPC server.
Public Const RPC_S_FP_OVERFLOW = 1771 'A floating-point overflow occurred at the RPC server.
Public Const RPC_X_NO_MORE_ENTRIES = 1772 'The list of RPC servers available for the binding of auto handles has been exhausted.
Public Const RPC_X_SS_CHAR_TRANS_OPEN_FAIL = 1773 'Unable to open the character translation table file.
Public Const RPC_X_SS_CHAR_TRANS_SHORT_FILE = 1774 'The file containing the character translation table has fewer than bytes.
Public Const RPC_X_SS_IN_NULL_CONTEXT = 1775 'A null context handle was passed from the client to the host during a remote procedure call.
Public Const RPC_X_SS_CONTEXT_DAMAGED = 1777 'The context handle changed during a remote procedure call.
Public Const RPC_X_SS_HANDLES_MISMATCH = 1778 'The binding handles passed to a remote procedure call do not match.
Public Const RPC_X_SS_CANNOT_GET_CALL_HANDLE = 1779 'The stub is unable to get the remote procedure call handle.
Public Const RPC_X_NULL_REF_POINTER = 1780 'A null reference pointer was passed to the stub.
Public Const RPC_X_ENUM_VALUE_OUT_OF_RANGE = 1781 'The enumeration value is out of range.
Public Const RPC_X_BYTE_COUNT_TOO_SMALL = 1782 'The byte count is too small.
Public Const RPC_X_BAD_STUB_DATA = 1783 'The stub received bad data.
Public Const ERROR_INVALID_USER_BUFFER = 1784 'The supplied user buffer is not valid for the requested operation.
Public Const ERROR_UNRECOGNIZED_MEDIA = 1785 'The disk media is not recognized. It may not be formatted.
Public Const ERROR_NO_TRUST_LSA_SECRET = 1786 'The workstation does not have a trust secret.
Public Const ERROR_NO_TRUST_SAM_ACCOUNT = 1787 'The SAM database on the Windows NT Server does not have a computer account for this workstation trust relationship.
Public Const ERROR_TRUSTED_DOMAIN_FAILURE = 1788 'The trust relationship between the primary domain and the trusted domain failed.
Public Const ERROR_TRUSTED_RELATIONSHIP_FAILURE = 1789 'The trust relationship between this workstation and the primary domain failed.
Public Const ERROR_TRUST_FAILURE = 1790 'The network logon failed.
Public Const RPC_S_CALL_IN_PROGRESS = 1791 'A remote procedure call is already in progress for this thread.
Public Const ERROR_NETLOGON_NOT_STARTED = 1792 'An attempt was made to logon, but the network logon service was not started.
Public Const ERROR_ACCOUNT_EXPIRED = 1793 'The user's account has expired.
Public Const ERROR_REDIRECTOR_HAS_OPEN_HANDLES = 1794 'The redirector is in use and cannot be unloaded.
Public Const ERROR_PRINTER_DRIVER_ALREADY_INSTALLED = 1795 'The specified printer driver is already installed.
Public Const ERROR_UNKNOWN_PORT = 1796 'The specified port is unknown.
Public Const ERROR_UNKNOWN_PRINTER_DRIVER = 1797 'The printer driver is unknown.
Public Const ERROR_UNKNOWN_PRINTPROCESSOR = 1798 'The print processor is unknown.
Public Const ERROR_INVALID_SEPARATOR_FILE = 1799 'The specified separator file is invalid.
Public Const ERROR_INVALID_PRIORITY = 1800 'The specified priority is invalid.
Public Const ERROR_INVALID_PRINTER_NAME = 1801 'The printer name is invalid.
Public Const ERROR_PRINTER_ALREADY_EXISTS = 1802 'The printer already exists.
Public Const ERROR_INVALID_PRINTER_COMMAND = 1803 'The printer command is invalid.
Public Const ERROR_INVALID_DATATYPE = 1804 'The specified datatype is invalid.
Public Const ERROR_INVALID_ENVIRONMENT = 1805 'The environment specified is invalid.
Public Const RPC_S_NO_MORE_BINDINGS = 1806 'There are no more bindings.
Public Const ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT = 1807 'The account used is an interdomain trust account. Use your global user account or local user account to access this server.
Public Const ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT = 1808 'The account used is a computer account. Use your global user account or local user account to access this server.
Public Const ERROR_NOLOGON_SERVER_TRUST_ACCOUNT = 1809 'The account used is a server trust account. Use your global user account or local user account to access this server.
Public Const ERROR_DOMAIN_TRUST_INCONSISTENT = 1810 'The name or security ID (SID) of the domain specified is inconsistent with the trust information for that domain.
Public Const ERROR_SERVER_HAS_OPEN_HANDLES = 1811 'The server is in use and cannot be unloaded.
Public Const ERROR_RESOURCE_DATA_NOT_FOUND = 1812 'The specified image file did not contain a resource section.
Public Const ERROR_RESOURCE_TYPE_NOT_FOUND = 1813 'The specified resource type cannot be found in the image file.
Public Const ERROR_RESOURCE_NAME_NOT_FOUND = 1814 'The specified resource name cannot be found in the image file.
Public Const ERROR_RESOURCE_LANG_NOT_FOUND = 1815 'The specified resource language ID cannot be found in the image file.
Public Const ERROR_NOT_ENOUGH_QUOTA = 1816 'Not enough quota is available to process this command.
Public Const RPC_S_NO_INTERFACES = 1817 'No interfaces have been registered.
Public Const RPC_S_CALL_CANCELLED = 1818 'The remote procedure call was cancelled.
Public Const RPC_S_BINDING_INCOMPLETE = 1819 'The binding handle does not contain all required information.
Public Const RPC_S_COMM_FAILURE = 1820 'A communications failure occurred during a remote procedure call.
Public Const RPC_S_UNSUPPORTED_AUTHN_LEVEL = 1821 'The requested authentication level is not supported.
Public Const RPC_S_NO_PRINC_NAME = 1822 'No principal name registered.
Public Const RPC_S_NOT_RPC_ERROR = 1823 'The error specified is not a valid Windows RPC error code.
Public Const RPC_S_UUID_LOCAL_ONLY = 1824 'A UUID that is valid only on this computer has been allocated.
Public Const RPC_S_SEC_PKG_ERROR = 1825 'A security package specific error occurred.
Public Const RPC_S_NOT_CANCELLED = 1826 'Thread is not canceled.
Public Const RPC_X_INVALID_ES_ACTION = 1827 'Invalid operation on the encoding/decoding handle.
Public Const RPC_X_WRONG_ES_VERSION = 1828 'Incompatible version of the serializing package.
Public Const RPC_X_WRONG_STUB_VERSION = 1829 'Incompatible version of the RPC stub.
Public Const RPC_X_INVALID_PIPE_OBJECT = 1830 'The RPC pipe object is invalid or corrupted.
Public Const RPC_X_WRONG_PIPE_ORDER = 1831 'An invalid operation was attempted on an RPC pipe object.
Public Const RPC_X_WRONG_PIPE_VERSION = 1832 'Unsupported RPC pipe version.
'* Error Codes 1833 - 1897 Not Listed
Public Const RPC_S_GROUP_MEMBER_NOT_FOUND = 1898 'The group member was not found.
Public Const EPT_S_CANT_CREATE = 1899 'The endpoint mapper database entry could not be created.
Public Const RPC_S_INVALID_OBJECT = 1900 'The object universal unique identifier (UUID) is the nil UUID.
Public Const ERROR_INVALID_TIME = 1901 'The specified time is invalid.
Public Const ERROR_INVALID_FORM_NAME = 1902 'The specified form name is invalid.
Public Const ERROR_INVALID_FORM_SIZE = 1903 'The specified form size is invalid.
Public Const ERROR_ALREADY_WAITING = 1904 'The specified printer handle is already being waited on.
Public Const ERROR_PRINTER_DELETED = 1905 'The specified printer has been deleted.
Public Const ERROR_INVALID_PRINTER_STATE = 1906 'The state of the printer is invalid.
Public Const ERROR_PASSWORD_MUST_CHANGE = 1907 'The user must change his password before he logs on the first time.
Public Const ERROR_DOMAIN_CONTROLLER_NOT_FOUND = 1908 'Could not find the domain controller for this domain.
Public Const ERROR_ACCOUNT_LOCKED_OUT = 1909 'The referenced account is currently locked out and may not be logged on to.
Public Const OR_INVALID_OXID = 1910 'The object exporter specified was not found.
Public Const OR_INVALID_OID = 1911 'The object specified was not found.
Public Const OR_INVALID_SET = 1912 'The object resolver set specified was not found.
Public Const RPC_S_SEND_INCOMPLETE = 1913 'Some data remains to be sent in the request buffer.
Public Const RPC_S_INVALID_ASYNC_HANDLE = 1914 'Invalid asynchronous remote procedure call handle.
Public Const RPC_S_INVALID_ASYNC_CALL = 1915 'Invalid asynchronous RPC call handle for this operation.
Public Const RPC_X_PIPE_CLOSED = 1916 'The RPC pipe object has already been closed.
Public Const RPC_X_PIPE_DISCIPLINE_ERROR = 1917 'The RPC call completed before all pipes were processed.
Public Const RPC_X_PIPE_EMPTY = 1918 'No more data is available from the RPC pipe.
Public Const ERROR_NO_SITENAME = 1919 'No site name is available for this machine.
Public Const ERROR_CANT_ACCESS_FILE = 1920 'The file can not be accessed by the system.
Public Const ERROR_CANT_RESOLVE_FILENAME = 1921 'The name of the file cannot be resolved by the system.
Public Const ERROR_DS_MEMBERSHIP_EVALUATED_LOCALLY = 1922 'The directory service evaluated group memberships locally.
Public Const ERROR_DS_NO_ATTRIBUTE_OR_VALUE = 1923 'The specified directory service attribute or value does not exist.
Public Const ERROR_DS_INVALID_ATTRIBUTE_SYNTAX = 1924 'The attribute syntax specified to the directory service is invalid.
Public Const ERROR_DS_ATTRIBUTE_TYPE_UNDEFINED = 1925 'The attribute type specified to the directory service is not defined.
Public Const ERROR_DS_ATTRIBUTE_OR_VALUE_EXISTS = 1926 'The specified directory service attribute or value already exists.
Public Const ERROR_DS_BUSY = 1927 'The directory service is busy.
Public Const ERROR_DS_UNAVAILABLE = 1928 'The directory service is unavailable.
Public Const ERROR_DS_NO_RIDS_ALLOCATED = 1929 'The directory service was unable to allocate a relative identifier.
Public Const ERROR_DS_NO_MORE_RIDS = 1930 'The directory service has exhausted the pool of relative identifiers.
Public Const ERROR_DS_INCORRECT_ROLE_OWNER = 1931 'The requested operation could not be performed because the directory service is not the master for that type of operation.
Public Const ERROR_DS_RIDMGR_INIT_ERROR = 1932 'The directory service was unable to initialize the subsystem that allocates relative identifiers.
Public Const ERROR_DS_OBJ_CLASS_VIOLATION = 1933 'The requested operation did not satisfy one or more constraints associated with the class of the object.
Public Const ERROR_DS_CANT_ON_NON_LEAF = 1934 'The directory service can perform the requested operation only on a leaf object.
Public Const ERROR_DS_CANT_ON_RDN = 1935 'The directory service cannot perform the requested operation on the RDN attribute of an object.
Public Const ERROR_DS_CANT_MOD_OBJ_CLASS = 1936 'The directory service detected an attempt to modify the object class of an object.
Public Const ERROR_DS_CROSS_DOM_MOVE_ERROR = 1937 'The requested cross domain move operation could not be performed.
Public Const ERROR_DS_GC_NOT_AVAILABLE = 1938 'Unable to contact the global catalog server.
'* Error Codes 1939 - 1999 Not Listed
Public Const ERROR_INVALID_PIXEL_FORMAT = 2000 'The pixel format is invalid.
Public Const ERROR_BAD_DRIVER = 2001 'The specified driver is invalid.
Public Const ERROR_INVALID_WINDOW_STYLE = 2002 'The window style or class attribute is invalid for this operation.
Public Const ERROR_METAFILE_NOT_SUPPORTED = 2003 'The requested metafile operation is not supported.
Public Const ERROR_TRANSFORM_NOT_SUPPORTED = 2004 'The requested transformation operation is not supported.
Public Const ERROR_CLIPPING_NOT_SUPPORTED = 2005 'The requested clipping operation is not supported.
'* Error Codes 2006 - 2107 Not Listed
Public Const ERROR_CONNECTED_OTHER_PASSWORD = 2108 'The network connection was made successfully, but the user had to be prompted for a password other than the one originally specified.
'* Error Codes 2109 - 2201 Not Listed
Public Const ERROR_BAD_USERNAME = 2202 'The specified username is invalid.
'* Error Codes 2203 - 2249 Not Listed
Public Const ERROR_NOT_CONNECTED = 2250 'This network connection does not exist.
'* Error Codes 2251 - 2299 Not Listed
Public Const ERROR_INVALID_CMM = 2300 'The specified color management module is invalid.
Public Const ERROR_INVALID_PROFILE = 2301 'The specified color profile is invalid.
Public Const ERROR_TAG_NOT_FOUND = 2302 'The specified tag was not found.
Public Const ERROR_TAG_NOT_PRESENT = 2303 'A required tag is not present.
Public Const ERROR_DUPLICATE_TAG = 2304 'The specified tag is already present.
Public Const ERROR_PROFILE_NOT_ASSOCIATED_WITH_DEVICE = 2305 'The specified color profile is not associated with any device.
Public Const ERROR_PROFILE_NOT_FOUND = 2306 'The specified color profile was not found.
Public Const ERROR_INVALID_COLORSPACE = 2307 'The specified color space is invalid.
Public Const ERROR_ICM_NOT_ENABLED = 2308 'Image Color Management is not enabled.
Public Const ERROR_DELETING_ICM_XFORM = 2309 'There was an error while deleting the color transform.
Public Const ERROR_INVALID_TRANSFORM = 2310 'The specified color transform is invalid.
'* Error Codes 2311 - 2400 Not Listed
Public Const ERROR_OPEN_FILES = 2401 'This network connection has files open or requests pending.
Public Const ERROR_ACTIVE_CONNECTIONS = 2402 'Active connections still exist.
'* Error Code 2403 Not Listed
Public Const ERROR_DEVICE_IN_USE = 2404 'The device is in use by an active process and cannot be disconnected.
'* Error Codes 2405 - 2999 Not Listed
Public Const ERROR_UNKNOWN_PRINT_MONITOR = 3000 'The specified print monitor is unknown.
Public Const ERROR_PRINTER_DRIVER_IN_USE = 3001 'The specified printer driver is currently in use.
Public Const ERROR_SPOOL_FILE_NOT_FOUND = 3002 'The spool file was not found.
Public Const ERROR_SPL_NO_STARTDOC = 3003 'A StartDocPrinter call was not issued.
Public Const ERROR_SPL_NO_ADDJOB = 3004 'An AddJob call was not issued.
Public Const ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED = 3005 'The specified print processor has already been installed.
Public Const ERROR_PRINT_MONITOR_ALREADY_INSTALLED = 3006 'The specified print monitor has already been installed.
Public Const ERROR_INVALID_PRINT_MONITOR = 3007 'The specified print monitor does not have the required functions.
Public Const ERROR_PRINT_MONITOR_IN_USE = 3008 'The specified print monitor is currently in use.
Public Const ERROR_PRINTER_HAS_JOBS_QUEUED = 3009 'The requested operation is not allowed when there are jobs queued to the printer.
Public Const ERROR_SUCCESS_REBOOT_REQUIRED = 3010 'The requested operation is successful. Changes will not be effective until the system is rebooted.
Public Const ERROR_SUCCESS_RESTART_REQUIRED = 3011 'The requested operation is successful. Changes will not be effective until the service is restarted.
'* Error Codes 3012 - 3999 Not Listed
Public Const ERROR_WINS_INTERNAL = 4000 'WINS encountered an error while processing the command.
Public Const ERROR_CAN_NOT_DEL_LOCAL_WINS = 4001 'The local WINS can not be deleted.
Public Const ERROR_STATIC_INIT = 4002 'The importation from the file failed.
Public Const ERROR_INC_BACKUP = 4003 'The backup failed. Was a full backup done before?
Public Const ERROR_FULL_BACKUP = 4004 'The backup failed. Check the directory to which you are backing the database.
Public Const ERROR_REC_NON_EXISTENT = 4005 'The name does not exist in the WINS database.
Public Const ERROR_RPL_NOT_ALLOWED = 4006 'Replication with a nonconfigured partner is not allowed.
'* Error Codes 4007 - 4099 Not Listed
Public Const ERROR_DHCP_ADDRESS_CONFLICT = 4100 'The DHCP client has obtained an IP address that is already in use on the network. The local interface will be disabled until the DHCP client can obtain a new address.
'* Error Codes 4101 - 4199 Not Listed
Public Const ERROR_WMI_GUID_NOT_FOUND = 4200 'The GUID passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_INSTANCE_NOT_FOUND = 4201 'The instance name passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_ITEMID_NOT_FOUND = 4202 'The data item ID passed was not recognized as valid by a WMI data provider.
Public Const ERROR_WMI_TRY_AGAIN = 4203 'The WMI request could not be completed and should be retried.
Public Const ERROR_WMI_DP_NOT_FOUND = 4204 'The WMI data provider could not be located.
Public Const ERROR_WMI_UNRESOLVED_INSTANCE_REF = 4205 'The WMI data provider references an instance set that has not been registered.
Public Const ERROR_WMI_ALREADY_ENABLED = 4206 'The WMI data block or event notification has already been enabled.
Public Const ERROR_WMI_GUID_DISCONNECTED = 4207 'The WMI data block is no longer available.
Public Const ERROR_WMI_SERVER_UNAVAILABLE = 4208 'The WMI data service is not available.
Public Const ERROR_WMI_DP_FAILED = 4209 'The WMI data provider failed to carry out the request.
Public Const ERROR_WMI_INVALID_MOF = 4210 'The WMI MOF information is not valid.
Public Const ERROR_WMI_INVALID_REGINFO = 4211 'The WMI registration information is not valid.
'* Error Codes 4212 - 4299 Not Listed
Public Const ERROR_INVALID_MEDIA = 4300 'The media identifier does not represent a valid medium.
Public Const ERROR_INVALID_LIBRARY = 4301 'The library identifier does not represent a valid library.
Public Const ERROR_INVALID_MEDIA_POOL = 4302 'The media pool identifier does not represent a valid media pool.
Public Const ERROR_DRIVE_MEDIA_MISMATCH = 4303 'The drive and medium are not compatible or exist in different libraries.
Public Const ERROR_MEDIA_OFFLINE = 4304 'The medium currently exists in an offline library and must be online to perform this operation.
Public Const ERROR_LIBRARY_OFFLINE = 4305 'The operation cannot be performed on an offline library.
Public Const ERROR_EMPTY = 4306 'The library, drive, or media pool is empty.
Public Const ERROR_NOT_EMPTY = 4307 'The library, drive, or media pool must be empty to perform this operation.
Public Const ERROR_MEDIA_UNAVAILABLE = 4308 'No media is currently available in this media pool or library.
Public Const ERROR_RESOURCE_DISABLED = 4309 'A resource required for this operation is disabled.
Public Const ERROR_INVALID_CLEANER = 4310 'The media identifier does not represent a valid cleaner.
Public Const ERROR_UNABLE_TO_CLEAN = 4311 'The drive cannot be cleaned or does not support cleaning.
Public Const ERROR_OBJECT_NOT_FOUND = 4312 'The object identifier does not represent a valid object.
Public Const ERROR_DATABASE_FAILURE = 4313 'Unable to read from or write to the database.
Public Const ERROR_DATABASE_FULL = 4314 'The database is full.
Public Const ERROR_MEDIA_INCOMPATIBLE = 4315 'The medium is not compatible with the device or media pool.
Public Const ERROR_RESOURCE_NOT_PRESENT = 4316 'The resource required for this operation does not exist.
Public Const ERROR_INVALID_OPERATION = 4317 'The operation identifier is not valid.
Public Const ERROR_MEDIA_NOT_AVAILABLE = 4318 'The media is not mounted or ready for use.
Public Const ERROR_DEVICE_NOT_AVAILABLE = 4319 'The device is not ready for use.
Public Const ERROR_REQUEST_REFUSED = 4320 'The operator or administrator has refused the request.
'* Error Codes 4321 - 4349 Not Listed
Public Const ERROR_FILE_OFFLINE = 4350 'The remote storage service was not able to recall the file.
Public Const ERROR_REMOTE_STORAGE_NOT_ACTIVE = 4351 'The remote storage service is not operational at this time.
Public Const ERROR_REMOTE_STORAGE_MEDIA_ERROR = 4352 'The remote storage service encountered a media error.
'* Error Codes 4353 - 4389 Not Listed
Public Const ERROR_NOT_A_REPARSE_POINT = 4390 'The file or directory is not a reparse point.
Public Const ERROR_REPARSE_ATTRIBUTE_CONFLICT = 4391 'The reparse point attribute cannot be set because it conflicts with an existing attribute.
'* Error Codes 4392 - 5000 Not Listed
Public Const ERROR_DEPENDENT_RESOURCE_EXISTS = 5001 'The cluster resource cannot be moved to another group because other resources are dependent on it.
Public Const ERROR_DEPENDENCY_NOT_FOUND = 5002 'The cluster resource dependency cannot be found.
Public Const ERROR_DEPENDENCY_ALREADY_EXISTS = 5003 'The cluster resource cannot be made dependent on the specified resource because it is already dependent.
Public Const ERROR_RESOURCE_NOT_ONLINE = 5004 'The cluster resource is not online.
Public Const ERROR_HOST_NODE_NOT_AVAILABLE = 5005 'A cluster node is not available for this operation.
Public Const ERROR_RESOURCE_NOT_AVAILABLE = 5006 'The cluster resource is not available.
Public Const ERROR_RESOURCE_NOT_FOUND = 5007 'The cluster resource could not be found.
Public Const ERROR_SHUTDOWN_CLUSTER = 5008 'The cluster is being shut down.
Public Const ERROR_CANT_EVICT_ACTIVE_NODE = 5009 'A cluster node cannot be evicted from the cluster while it is online.
Public Const ERROR_OBJECT_ALREADY_EXISTS = 5010 'The object already exists.
Public Const ERROR_OBJECT_IN_LIST = 5011 'The object is already in the list.
Public Const ERROR_GROUP_NOT_AVAILABLE = 5012 'The cluster group is not available for any new requests.
Public Const ERROR_GROUP_NOT_FOUND = 5013 'The cluster group could not be found.
Public Const ERROR_GROUP_NOT_ONLINE = 5014 'The operation could not be completed because the cluster group is not online.
Public Const ERROR_HOST_NODE_NOT_RESOURCE_OWNER = 5015 'The cluster node is not the owner of the resource.
Public Const ERROR_HOST_NODE_NOT_GROUP_OWNER = 5016 'The cluster node is not the owner of the group.
Public Const ERROR_RESMON_CREATE_FAILED = 5017 'The cluster resource could not be created in the specified resource monitor.
Public Const ERROR_RESMON_ONLINE_FAILED = 5018 'The cluster resource could not be brought online by the resource monitor.
Public Const ERROR_RESOURCE_ONLINE = 5019 'The operation could not be completed because the cluster resource is online.
Public Const ERROR_QUORUM_RESOURCE = 5020 'The cluster resource could not be deleted or brought offline because it is the quorum resource.
Public Const ERROR_NOT_QUORUM_CAPABLE = 5021 'The cluster could not make the specified resource a quorum resource because it is not capable of being a quorum resource.
Public Const ERROR_CLUSTER_SHUTTING_DOWN = 5022 'The cluster software is shutting down.
Public Const ERROR_INVALID_STATE = 5023 'The group or resource is not in the correct state to perform the requested operation.  ERROR_INVALID_STATE
Public Const ERROR_RESOURCE_PROPERTIES_STORED = 5024 'The properties were stored but not all changes will take effect until the next time the resource is brought online.
Public Const ERROR_NOT_QUORUM_CLASS = 5025 'The cluster could not make the specified resource a quorum resource because it does not belong to a shared storage class.
Public Const ERROR_CORE_RESOURCE = 5026 'The cluster resource could not be deleted since it is a core resource.
Public Const ERROR_QUORUM_RESOURCE_ONLINE_FAILED = 5027 'The quorum resource failed to come online.
Public Const ERROR_QUORUMLOG_OPEN_FAILED = 5028 'The quorum log could not be created or mounted successfully.
Public Const ERROR_CLUSTERLOG_CORRUPT = 5029 'The cluster log is corrupt.
Public Const ERROR_CLUSTERLOG_RECORD_EXCEEDS_MAXSIZE = 5030 'The record could not be written to the cluster log since it exceeds the maximum size.
Public Const ERROR_CLUSTERLOG_EXCEEDS_MAXSIZE = 5031 'The cluster log exceeds its maximum size.
Public Const ERROR_CLUSTERLOG_CHKPOINT_NOT_FOUND = 5032 'No checkpoint record was found in the cluster log.
Public Const ERROR_CLUSTERLOG_NOT_ENOUGH_SPACE = 5033 'The minimum required disk space needed for logging is not available.
'* Error Codes 5034 - 5999 Not Listed
Public Const ERROR_ENCRYPTION_FAILED = 6000 'The specified file could not be encrypted.
Public Const ERROR_DECRYPTION_FAILED = 6001 'The specified file could not be decrypted.
Public Const ERROR_FILE_ENCRYPTED = 6002 'The specified file is encrypted and the user does not have the ability to decrypt it.
Public Const ERROR_NO_RECOVERY_POLICY = 6003 'There is no encryption recovery policy configured for this system.
Public Const ERROR_NO_EFS = 6004 'The required encryption driver is not loaded for this system.
Public Const ERROR_WRONG_EFS = 6005 'The file was encrypted with a different encryption driver than is currently loaded.
Public Const ERROR_NO_USER_KEYS = 6006 'There are no EFS keys defined for the user.
Public Const ERROR_FILE_NOT_ENCRYPTED = 6007 'The specified file is not encrypted.
Public Const ERROR_NOT_EXPORT_FORMAT = 6008 'The specified file is not in the defined EFS export format.
'* Error Codes 6009 - 6117 Not Listed
Public Const ERROR_NO_BROWSER_SERVERS_FOUND = 6118 'The list of servers for this workgroup is not currently available

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias _
      "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, _
      ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
      ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) _
      As Long
Public Function GetErrorText(lCode As Long) As String
    Dim sRtrnCode As String
    Dim lRet As Long
    sRtrnCode = Space$(256)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lCode, 0&, sRtrnCode, 256&, 0&)
    If lRet > 0 Then
       GetErrorText = Left(sRtrnCode, lRet)
    Else
        GetErrorText = "Error text unavailable"
    End If
End Function
Public Function GetErrorTitle(inErrNum As Integer) As String
    '********************************************************************
    '* This function receives the error number and returns the
    '* error title.
    '********************************************************************
    Dim stTitle As String
    
    '*******************************************************************
    '* Compare the error number to ranges of integers to get to the
    '* correct range, then use Select Case to get the title
    '*******************************************************************
    '* Remember that Error Code 0 is success, you should handle this
    '* in your code or you may get error handling for successful calls
    '*******************************************************************
    If inErrNum = 0 Then
        '* Success
        stTitle = "ERROR_SUCCESS"
    ElseIf inErrNum < 0 Then
        stTitle = "Error Number Out Of Range (Less Than 0)"
    ElseIf inErrNum > 0 And inErrNum <= 200 Then
        '* Handle Range 1 - 200
        If inErrNum > 0 And inErrNum <= 100 Then
            '* Handle Range 1 to 100
            Select Case inErrNum
            Case 1
                stTitle = "ERROR_INVALID_FUNCTION"
            Case 2
                stTitle = "ERROR_FILE_NOT_FOUND"
            Case 3
                stTitle = "ERROR_PATH_NOT_FOUND"
            Case 4
                stTitle = "ERROR_TOO_MANY_OPEN_FILES"
            Case 5
                stTitle = "ERROR_ACCESS_DENIED"
            Case 6
                stTitle = "ERROR_INVALID_HANDLE"
            Case 7
                stTitle = "ERROR_ARENA_TRASHED"
            Case 8
                stTitle = "ERROR_NOT_ENOUGH_MEMORY"
            Case 9
                stTitle = "ERROR_INVALID_BLOCK"
            Case 10
                stTitle = "ERROR_BAD_ENVIRONMENT"
            Case 11
                stTitle = "ERROR_BAD_FORMAT"
            Case 12
                stTitle = "ERROR_INVALID_ACCESS"
            Case 13
                stTitle = "ERROR_INVALID_DATA"
            Case 14
                stTitle = "ERROR_OUTOFMEMORY"
            Case 15
                stTitle = "ERROR_INVALID_DRIVE"
            Case 16
                stTitle = "ERROR_CURRENT_DIRECTORY"
            Case 17
                stTitle = "ERROR_NOT_SAME_DEVICE"
            Case 18
                stTitle = "ERROR_NO_MORE_FILES"
            Case 19
                stTitle = "ERROR_WRITE_PROTECT"
            Case 20
                stTitle = "ERROR_BAD_UNIT"
            Case 21
                stTitle = "ERROR_NOT_READY"
            Case 22
                stTitle = "ERROR_BAD_COMMAND"
            Case 23
                stTitle = "ERROR_CRC"
            Case 24
                stTitle = "ERROR_BAD_LENGTH"
            Case 25
                stTitle = "ERROR_SEEK"
'************************************************
'* Again, the code gets boring here
'* <CLIP>
'************************************************
            Case 26
                stTitle = "ERROR_NOT_DOS_DISK"
            Case 27
                stTitle = "ERROR_SECTOR_NOT_FOUND"
            Case 28
                stTitle = "ERROR_OUT_OF_PAPER"
            Case 29
                stTitle = "ERROR_WRITE_FAULT"
            Case 30
                stTitle = "ERROR_READ_FAULT"
            Case 31
                stTitle = "ERROR_GEN_FAILURE"
            Case 32
                stTitle = "ERROR_SHARING_VIOLATION"
            Case 33
                stTitle = "ERROR_LOCK_VIOLATION"
            Case 34
                stTitle = "ERROR_WRONG_DISK"
            Case 35
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 36
                stTitle = "ERROR_SHARING_BUFFER_EXCEEDED"
            Case 38
                stTitle = "ERROR_HANDLE_EOF"
            Case 39
                stTitle = "ERROR_HANDLE_DISK_FULL"
            Case 40 To 49
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 50
                stTitle = "ERROR_NOT_SUPPORTED"
            Case 51
                stTitle = "ERROR_REM_NOT_LIST"
            Case 52
                stTitle = "ERROR_DUP_NAME"
            Case 53
                stTitle = "ERROR_BAD_NETPATH"
            Case 54
                stTitle = "ERROR_NETWORK_BUSY"
            Case 55
                stTitle = "ERROR_DEV_NOT_EXIST"
            Case 56
                stTitle = "ERROR_TOO_MANY_CMDS"
            Case 57
                stTitle = "ERROR_ADAP_HDW_ERR"
            Case 58
                stTitle = "ERROR_BAD_NET_RESP"
            Case 59
                stTitle = "ERROR_UNEXP_NET_ERR"
            Case 60
                stTitle = "ERROR_BAD_REM_ADAP"
            Case 61
                stTitle = "ERROR_PRINTQ_FULL"
            Case 62
                stTitle = "ERROR_NO_SPOOL_SPACE"
            Case 63
                stTitle = "ERROR_PRINT_CANCELLED"
            Case 64
                stTitle = "ERROR_NETNAME_DELETED"
            Case 65
                stTitle = "ERROR_NETWORK_ACCESS_DENIED"
            Case 66
                stTitle = "ERROR_BAD_DEV_TYPE"
            Case 67
                stTitle = "ERROR_BAD_NET_NAME"
            Case 68
                stTitle = "ERROR_TOO_MANY_NAMES"
            Case 69
                stTitle = "ERROR_TOO_MANY_SESS"
            Case 70
                stTitle = "ERROR_SHARING_PAUSED"
            Case 71
                stTitle = "ERROR_REQ_NOT_ACCEP"
            Case 72
                stTitle = "ERROR_REDIR_PAUSED"
            Case 73 To 79
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 80
                stTitle = "ERROR_FILE_EXISTS"
            Case 81
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 82
                stTitle = "ERROR_CANNOT_MAKE"
            Case 83
                stTitle = "ERROR_FAIL_I24"
            Case 84
                stTitle = "ERROR_OUT_OF_STRUCTURES"
            Case 85
                stTitle = "ERROR_ALREADY_ASSIGNED"
            Case 86
                stTitle = "ERROR_INVALID_PASSWORD"
            Case 87
                stTitle = "ERROR_INVALID_PARAMETER"
            Case 88
                stTitle = "ERROR_NET_WRITE_FAULT"
            Case 89
                stTitle = "ERROR_NO_PROC_SLOTS"
            Case 90 To 99
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 100
                stTitle = "ERROR_TOO_MANY_SEMAPHORES"
            End Select
        ElseIf inErrNum > 100 And inErrNum <= 200 Then
            '* Handle Range 101 to 200
            Select Case inErrNum
            Case 101
                stTitle = "ERROR_EXCL_SEM_ALREADY_OWNED"
            Case 102
                stTitle = "ERROR_SEM_IS_SET"
            Case 103
                stTitle = "ERROR_TOO_MANY_SEM_REQUESTS"
            Case 104
                stTitle = "ERROR_INVALID_AT_INTERRUPT_TIME"
            Case 105
                stTitle = "ERROR_SEM_OWNER_DIED"
            Case 106
                stTitle = "ERROR_SEM_USER_LIMIT"
            Case 107
                stTitle = "ERROR_DISK_CHANGE"
            Case 108
                stTitle = "ERROR_DRIVE_LOCKED"
            Case 109
                stTitle = "ERROR_BROKEN_PIPE"
            Case 110
                stTitle = "ERROR_OPEN_FAILED"
            Case 111
                stTitle = "ERROR_BUFFER_OVERFLOW"
            Case 112
                stTitle = "ERROR_DISK_FULL"
            Case 113
                stTitle = "ERROR_NO_MORE_SEARCH_HANDLES"
            Case 114
                stTitle = "ERROR_INVALID_TARGET_HANDLE"
            Case 115 To 116
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 117
                stTitle = "ERROR_INVALID_CATEGORY"
            Case 118
                stTitle = "ERROR_INVALID_VERIFY_SWITCH"
            Case 119
                stTitle = "ERROR_BAD_DRIVER_LEVEL"
            Case 120
                stTitle = "ERROR_CALL_NOT_IMPLEMENTED"
            Case 121
                stTitle = "ERROR_SEM_TIMEOUT"
            Case 122
                stTitle = "ERROR_INSUFFICIENT_BUFFER"
            Case 123
                stTitle = "ERROR_INVALID_NAME"
            Case 124
                stTitle = "ERROR_INVALID_LEVEL"
            Case 125
                stTitle = "ERROR_NO_VOLUME_LABEL"
            Case 126
                stTitle = "ERROR_MOD_NOT_FOUND"
            Case 127
                stTitle = "ERROR_PROC_NOT_FOUND"
            Case 128
                stTitle = "ERROR_WAIT_NO_CHILDREN"
            Case 129
                stTitle = "ERROR_CHILD_NOT_COMPLETE"
            Case 130
                stTitle = "ERROR_DIRECT_ACCESS_HANDLE"
            Case 131
                stTitle = "ERROR_NEGATIVE_SEEK"
            Case 132
                stTitle = "ERROR_SEEK_ON_DEVICE"
            Case 133
                stTitle = "ERROR_IS_JOIN_TARGET"
            Case 134
                stTitle = "ERROR_IS_JOINED"
            Case 135
                stTitle = "ERROR_IS_SUBSTED"
            Case 136
                stTitle = "ERROR_NOT_JOINED"
            Case 137
                stTitle = "ERROR_NOT_SUBSTED"
            Case 138
                stTitle = "ERROR_JOIN_TO_JOIN"
            Case 139
                stTitle = "ERROR_SUBST_TO_SUBST"
            Case 140
                stTitle = "ERROR_JOIN_TO_SUBST"
            Case 141
                stTitle = "ERROR_SUBST_TO_JOIN"
            Case 142
                stTitle = "ERROR_BUSY_DRIVE"
            Case 143
                stTitle = "ERROR_SAME_DRIVE"
            Case 144
                stTitle = "ERROR_DIR_NOT_ROOT"
            Case 145
                stTitle = "ERROR_DIR_NOT_EMPTY"
            Case 146
                stTitle = "ERROR_IS_SUBST_PATH"
            Case 147
                stTitle = "ERROR_IS_JOIN_PATH"
            Case 148
                stTitle = "ERROR_PATH_BUSY"
            Case 149
                stTitle = "ERROR_IS_SUBST_TARGET"
            Case 150
                stTitle = "ERROR_SYSTEM_TRACE"
            Case 151
                stTitle = "ERROR_INVALID_EVENT_COUNT"
            Case 152
                stTitle = "ERROR_TOO_MANY_MUXWAITERS"
            Case 153
                stTitle = "ERROR_INVALID_LIST_FORMAT"
            Case 154
                stTitle = "ERROR_LABEL_TOO_LONG"
            Case 155
                stTitle = "ERROR_TOO_MANY_TCBS"
            Case 156
                stTitle = "ERROR_SIGNAL_REFUSED"
            Case 157
                stTitle = "ERROR_DISCARDED"
            Case 158
                stTitle = "ERROR_NOT_LOCKED"
            Case 159
                stTitle = "ERROR_BAD_THREADID_ADDR"
            Case 160
                stTitle = "ERROR_BAD_ARGUMENTS"
            Case 161
                stTitle = "ERROR_BAD_PATHNAME"
            Case 162
                stTitle = "ERROR_SIGNAL_PENDING"
            Case 163
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 164
                stTitle = "ERROR_MAX_THRDS_REACHED"
            Case 165 To 166
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 167
                stTitle = "ERROR_LOCK_FAILED"
            Case 168 To 169
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 170
                stTitle = "ERROR_BUSY"
            Case 171 To 172
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 173
                stTitle = "ERROR_CANCEL_VIOLATION"
            Case 174
                stTitle = "ERROR_ATOMIC_LOCKS_NOT_SUPPORTED"
            Case 175 To 179
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 180
                stTitle = "ERROR_INVALID_SEGMENT_NUMBER"
            Case 181
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 182
                stTitle = "ERROR_INVALID_ORDINAL"
            Case 183
                stTitle = "ERROR_ALREADY_EXISTS"
            Case 184 To 185
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 186
                stTitle = "ERROR_INVALID_FLAG_NUMBER"
            Case 187
                stTitle = "ERROR_SEM_NOT_FOUND"
            Case 188
                stTitle = "ERROR_INVALID_STARTING_CODESEG"
            Case 189
                stTitle = "ERROR_INVALID_STACKSEG"
            Case 190
                stTitle = "ERROR_INVALID_MODULETYPE"
            Case 191
                stTitle = "ERROR_INVALID_EXE_SIGNATURE"
            Case 192
                stTitle = "ERROR_EXE_MARKED_INVALID"
            Case 193
                stTitle = "ERROR_BAD_EXE_FORMAT"
            Case 194
                stTitle = "ERROR_ITERATED_DATA_EXCEEDS_64k"
            Case 195
                stTitle = "ERROR_INVALID_MINALLOCSIZE"
            Case 196
                stTitle = "ERROR_DYNLINK_FROM_INVALID_RING"
            Case 197
                stTitle = "ERROR_IOPL_NOT_ENABLED"
            Case 198
                stTitle = "ERROR_INVALID_SEGDPL"
            Case 199
                stTitle = "ERROR_AUTODATASEG_EXCEEDS_64k"
            Case 200
                stTitle = "ERROR_RING2SEG_MUST_BE_MOVABLE"
            End Select
        End If
    ElseIf inErrNum > 200 And inErrNum <= 1400 Then
        '* Handle Range 201 to 1400
        If inErrNum > 200 And inErrNum <= 1000 Then
            '* Handle Range 201 to 1000
            Select Case inErrNum
            Case 201
                stTitle = "ERROR_RELOC_CHAIN_XEEDS_SEGLIM"
            Case 202
                stTitle = "ERROR_INFLOOP_IN_RELOC_CHAIN"
            Case 203
                stTitle = "ERROR_ENVVAR_NOT_FOUND"
            Case 204
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 205
                stTitle = "ERROR_NO_SIGNAL_SENT"
            Case 206
                stTitle = "ERROR_FILENAME_EXCED_RANGE"
            Case 207
                stTitle = "ERROR_RING2_STACK_IN_USE"
            Case 208
                stTitle = "ERROR_META_EXPANSION_TOO_LONG"
            Case 209
                stTitle = "ERROR_INVALID_SIGNAL_NUMBER"
            Case 210
                stTitle = "ERROR_THREAD_1_INACTIVE"
            Case 211
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 212
                stTitle = "ERROR_LOCKED"
            Case 213
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 214
                stTitle = "ERROR_TOO_MANY_MODULES"
            Case 215
                stTitle = "ERROR_NESTING_NOT_ALLOWED"
            Case 216
                stTitle = "ERROR_EXE_MACHINE_TYPE_MISMATCH"
            Case 217 To 229
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 230
                stTitle = "ERROR_BAD_PIPE"
            Case 231
                stTitle = "ERROR_PIPE_BUSY"
            Case 232
                stTitle = "ERROR_NO_DATA"
            Case 233
                stTitle = "ERROR_PIPE_NOT_CONNECTED"
            Case 234
                stTitle = "ERROR_MORE_DATA"
            Case 235 To 239
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 240
                stTitle = "ERROR_VC_DISCONNECTED"
            Case 241 To 253
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 254
                stTitle = "ERROR_INVALID_EA_NAME"
            Case 255
                stTitle = "ERROR_EA_LIST_INCONSISTENT"
            Case 256 To 258
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 259
                stTitle = "ERROR_NO_MORE_ITEMS"
            Case 260 To 265
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 266
                stTitle = "ERROR_CANNOT_COPY"
            Case 267
                stTitle = "ERROR_DIRECTORY"
            Case 268 To 274
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 275
                stTitle = "ERROR_EAS_DIDNT_FIT"
            Case 276
                stTitle = "ERROR_EA_FILE_CORRUPT"
            Case 277
                stTitle = "ERROR_EA_TABLE_FULL"
            Case 278
                stTitle = "ERROR_INVALID_EA_HANDLE"
            Case 279 To 281
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 282
                stTitle = "ERROR_EAS_NOT_SUPPORTED"
            Case 283 To 287
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 288
                stTitle = "ERROR_NOT_OWNER"
            Case 289 To 297
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 298
                stTitle = "ERROR_TOO_MANY_POSTS"
            Case 299
                stTitle = "ERROR_PARTIAL_COPY"
            Case 300
                stTitle = "ERROR_OPLOCK_NOT_GRANTED"
            Case 301
                stTitle = "ERROR_INVALID_OPLOCK_PROTOCOL"
            Case 302 To 316
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 317
                stTitle = "ERROR_MR_MID_NOT_FOUND"
            Case 318 To 486
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 487
                stTitle = "ERROR_INVALID_ADDRESS"
            Case 488 To 533
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 534
                stTitle = "ERROR_ARITHMETIC_OVERFLOW"
            Case 535
                stTitle = "ERROR_PIPE_CONNECTED"
            Case 536
                stTitle = "ERROR_PIPE_LISTENING"
            Case 537 To 993
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 994
                stTitle = "ERROR_EA_ACCESS_DENIED"
            Case 995
                stTitle = "ERROR_OPERATION_ABORTED"
            Case 996
                stTitle = "ERROR_IO_INCOMPLETE"
            Case 997
                stTitle = "ERROR_IO_PENDING"
            Case 998
                stTitle = "ERROR_NOACCESS"
            Case 999
                stTitle = "ERROR_SWAPERROR"
            Case 1000
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 1000 And inErrNum <= 1200 Then
            '* Handle Range 1001 to 1200
            Select Case inErrNum
            Case 1001
                stTitle = "ERROR_STACK_OVERFLOW"
            Case 1002
                stTitle = "ERROR_INVALID_MESSAGE"
            Case 1003
                stTitle = "ERROR_CAN_NOT_COMPLETE"
            Case 1004
                stTitle = "ERROR_INVALID_FLAGS"
            Case 1005
                stTitle = "ERROR_UNRECOGNIZED_VOLUME"
            Case 1006
                stTitle = "ERROR_FILE_INVALID"
            Case 1007
                stTitle = "ERROR_FULLSCREEN_MODE"
            Case 1008
                stTitle = "ERROR_NO_TOKEN"
            Case 1009
                stTitle = "ERROR_BADDB"
            Case 1010
                stTitle = "ERROR_BADKEY"
            Case 1011
                stTitle = "ERROR_CANTOPEN"
            Case 1012
                stTitle = "ERROR_CANTREAD"
            Case 1013
                stTitle = "ERROR_CANTWRITE"
            Case 1014
                stTitle = "ERROR_REGISTRY_RECOVERED"
            Case 1015
                stTitle = "ERROR_REGISTRY_CORRUPT"
            Case 1016
                stTitle = "ERROR_REGISTRY_IO_FAILED"
            Case 1017
                stTitle = "ERROR_NOT_REGISTRY_FILE"
            Case 1018
                stTitle = "ERROR_KEY_DELETED"
            Case 1019
                stTitle = "ERROR_NO_LOG_SPACE"
            Case 1020
                stTitle = "ERROR_KEY_HAS_CHILDREN"
            Case 1021
                stTitle = "ERROR_CHILD_MUST_BE_VOLATILE"
            Case 1022
                stTitle = "ERROR_NOTIFY_ENUM_DIR"
            Case 1023 To 1050
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1051
                stTitle = "ERROR_DEPENDENT_SERVICES_RUNNING"
            Case 1052
                stTitle = "ERROR_INVALID_SERVICE_CONTROL"
            Case 1053
                stTitle = "ERROR_SERVICE_REQUEST_TIMEOUT"
            Case 1054
                stTitle = "ERROR_SERVICE_NO_THREAD"
            Case 1055
                stTitle = "ERROR_SERVICE_DATABASE_LOCKED"
            Case 1056
                stTitle = "ERROR_SERVICE_ALREADY_RUNNING"
            Case 1057
                stTitle = "ERROR_INVALID_SERVICE_ACCOUNT"
            Case 1058
                stTitle = "ERROR_SERVICE_DISABLED"
            Case 1059
                stTitle = "ERROR_CIRCULAR_DEPENDENCY"
            Case 1060
                stTitle = "ERROR_SERVICE_DOES_NOT_EXIST"
            Case 1061
                stTitle = "ERROR_SERVICE_CANNOT_ACCEPT_CTRL"
            Case 1062
                stTitle = "ERROR_SERVICE_NOT_ACTIVE"
            Case 1063
                stTitle = "ERROR_FAILED_SERVICE_CONTROLLER_CONNECT"
            Case 1064
                stTitle = "ERROR_EXCEPTION_IN_SERVICE"
            Case 1065
                stTitle = "ERROR_DATABASE_DOES_NOT_EXIST"
            Case 1066
                stTitle = "ERROR_SERVICE_SPECIFIC_ERROR"
            Case 1067
                stTitle = "ERROR_PROCESS_ABORTED"
            Case 1068
                stTitle = "ERROR_SERVICE_DEPENDENCY_FAIL"
            Case 1069
                stTitle = "ERROR_SERVICE_LOGON_FAILED"
            Case 1070
                stTitle = "ERROR_SERVICE_START_HANG"
            Case 1071
                stTitle = "ERROR_INVALID_SERVICE_LOCK"
            Case 1072
                stTitle = "ERROR_SERVICE_MARKED_FOR_DELETE"
            Case 1073
                stTitle = "ERROR_SERVICE_EXISTS"
            Case 1074
                stTitle = "ERROR_ALREADY_RUNNING_LKG"
            Case 1075
                stTitle = "ERROR_SERVICE_DEPENDENCY_DELETED"
            Case 1076
                stTitle = "ERROR_BOOT_ALREADY_ACCEPTED"
            Case 1077
                stTitle = "ERROR_SERVICE_NEVER_STARTED"
            Case 1078
                stTitle = "ERROR_DUPLICATE_SERVICE_NAME"
            Case 1079
                stTitle = "ERROR_DIFFERENT_SERVICE_ACCOUNT"
            Case 1080
                stTitle = "ERROR_CANNOT_DETECT_DRIVER_FAILURE"
            Case 1081
                stTitle = "ERROR_CANNOT_DETECT_PROCESS_ABORT"
            Case 1082
                stTitle = "ERROR_NO_RECOVERY_PROGRAM"
            Case 1083 To 1099
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1100
                stTitle = "ERROR_END_OF_MEDIA"
            Case 1101
                stTitle = "ERROR_FILEMARK_DETECTED"
            Case 1102
                stTitle = "ERROR_BEGINNING_OF_MEDIA"
            Case 1103
                stTitle = "ERROR_SETMARK_DETECTED"
            Case 1104
                stTitle = "ERROR_NO_DATA_DETECTED"
            Case 1105
                stTitle = "ERROR_PARTITION_FAILURE"
            Case 1106
                stTitle = "ERROR_INVALID_BLOCK_LENGTH"
            Case 1107
                stTitle = "ERROR_DEVICE_NOT_PARTITIONED"
            Case 1108
                stTitle = "ERROR_UNABLE_TO_LOCK_MEDIA"
            Case 1109
                stTitle = "ERROR_UNABLE_TO_UNLOAD_MEDIA"
            Case 1110
                stTitle = "ERROR_MEDIA_CHANGED"
            Case 1111
                stTitle = "ERROR_BUS_RESET"
            Case 1112
                stTitle = "ERROR_NO_MEDIA_IN_DRIVE"
            Case 1113
                stTitle = "ERROR_NO_UNICODE_TRANSLATION"
            Case 1114
                stTitle = "ERROR_DLL_INIT_FAILED"
            Case 1115
                stTitle = "ERROR_SHUTDOWN_IN_PROGRESS"
            Case 1116
                stTitle = "ERROR_NO_SHUTDOWN_IN_PROGRESS"
            Case 1117
                stTitle = "ERROR_IO_DEVICE"
            Case 1118
                stTitle = "ERROR_SERIAL_NO_DEVICE"
            Case 1119
                stTitle = "ERROR_IRQ_BUSY"
            Case 1120
                stTitle = "ERROR_MORE_WRITES"
            Case 1121
                stTitle = "ERROR_COUNTER_TIMEOUT"
            Case 1122
                stTitle = "ERROR_FLOPPY_ID_MARK_NOT_FOUND"
            Case 1123
                stTitle = "ERROR_FLOPPY_WRONG_CYLINDER"
            Case 1124
                stTitle = "ERROR_FLOPPY_UNKNOWN_ERROR"
            Case 1125
                stTitle = "ERROR_FLOPPY_BAD_REGISTERS"
            Case 1126
                stTitle = "ERROR_DISK_RECALIBRATE_FAILED"
            Case 1127
                stTitle = "ERROR_DISK_OPERATION_FAILED"
            Case 1128
                stTitle = "ERROR_DISK_RESET_FAILED"
            Case 1129
                stTitle = "ERROR_EOM_OVERFLOW"
            Case 1130
                stTitle = "ERROR_NOT_ENOUGH_SERVER_MEMORY"
            Case 1131
                stTitle = "ERROR_POSSIBLE_DEADLOCK"
            Case 1132
                stTitle = "ERROR_MAPPED_ALIGNMENT"
            Case 1133 To 1139
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1140
                stTitle = "ERROR_SET_POWER_STATE_VETOED"
            Case 1141
                stTitle = "ERROR_SET_POWER_STATE_FAILED"
            Case 1142
                stTitle = "ERROR_TOO_MANY_LINKS"
            Case 1143 To 1149
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1150
                stTitle = "ERROR_OLD_WIN_VERSION"
            Case 1151
                stTitle = "ERROR_APP_WRONG_OS"
            Case 1152
                stTitle = "ERROR_SINGLE_INSTANCE_APP"
            Case 1153
                stTitle = "ERROR_RMODE_APP"
            Case 1154
                stTitle = "ERROR_INVALID_DLL"
            Case 1155
                stTitle = "ERROR_NO_ASSOCIATION"
            Case 1156
                stTitle = "ERROR_DDE_FAIL"
            Case 1157
                stTitle = "ERROR_DLL_NOT_FOUND"
            Case 1158
                stTitle = "ERROR_NO_MORE_USER_HANDLES"
            Case 1159
                stTitle = "ERROR_MESSAGE_SYNC_ONLY"
            Case 1160
                stTitle = "ERROR_SOURCE_ELEMENT_EMPTY"
            Case 1161
                stTitle = "ERROR_DESTINATION_ELEMENT_FULL"
            Case 1162
                stTitle = "ERROR_ILLEGAL_ELEMENT_ADDRESS"
            Case 1163
                stTitle = "ERROR_MAGAZINE_NOT_PRESENT"
            Case 1164
                stTitle = "ERROR_DEVICE_REINITIALIZATION_NEEDED"
            Case 1165
                stTitle = "ERROR_DEVICE_REQUIRES_CLEANING"
            Case 1166
                stTitle = "ERROR_DEVICE_DOOR_OPEN"
            Case 1167
                stTitle = "ERROR_DEVICE_NOT_CONNECTED"
            Case 1168
                stTitle = "ERROR_NOT_FOUND"
            Case 1169
                stTitle = "ERROR_NO_MATCH"
            Case 1170
                stTitle = "ERROR_SET_NOT_FOUND"
            Case 1171
                stTitle = "ERROR_POINT_NOT_FOUND"
            Case 1172
                stTitle = "ERROR_NO_TRACKING_SERVICE"
            Case 1173
                stTitle = "ERROR_NO_VOLUME_ID"
            Case 1174 To 1199
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1200
                stTitle = "ERROR_BAD_DEVICE"
            End Select
        ElseIf inErrNum > 1200 And inErrNum <= 1400 Then
            '* Handle Range 1201 to 1400
            Select Case inErrNum
            Case 1201
                stTitle = "ERROR_CONNECTION_UNAVAIL"
            Case 1202
                stTitle = "ERROR_DEVICE_ALREADY_REMEMBERED"
            Case 1203
                stTitle = "ERROR_NO_NET_OR_BAD_PATH"
            Case 1204
                stTitle = "ERROR_BAD_PROVIDER"
            Case 1205
                stTitle = "ERROR_CANNOT_OPEN_PROFILE"
            Case 1206
                stTitle = "ERROR_BAD_PROFILE"
            Case 1207
                stTitle = "ERROR_NOT_CONTAINER"
            Case 1208
                stTitle = "ERROR_EXTENDED_ERROR"
            Case 1209
                stTitle = "ERROR_INVALID_GROUPNAME"
            Case 1210
                stTitle = "ERROR_INVALID_COMPUTERNAME"
            Case 1211
                stTitle = "ERROR_INVALID_EVENTNAME"
            Case 1212
                stTitle = "ERROR_INVALID_DOMAINNAME"
            Case 1213
                stTitle = "ERROR_INVALID_SERVICENAME"
            Case 1214
                stTitle = "ERROR_INVALID_NETNAME"
            Case 1215
                stTitle = "ERROR_INVALID_SHARENAME"
            Case 1216
                stTitle = "ERROR_INVALID_PASSWORDNAME"
            Case 1217
                stTitle = "ERROR_INVALID_MESSAGENAME"
            Case 1218
                stTitle = "ERROR_INVALID_MESSAGEDEST"
            Case 1219
                stTitle = "ERROR_SESSION_CREDENTIAL_CONFLICT"
            Case 1220
                stTitle = "ERROR_REMOTE_SESSION_LIMIT_EXCEEDED"
            Case 1221
                stTitle = "ERROR_DUP_DOMAINNAME"
            Case 1222
                stTitle = "ERROR_NO_NETWORK"
            Case 1223
                stTitle = "ERROR_CANCELLED"
            Case 1224
                stTitle = "ERROR_USER_MAPPED_FILE"
            Case 1225
                stTitle = "ERROR_CONNECTION_REFUSED"
            Case 1226
                stTitle = "ERROR_GRACEFUL_DISCONNECT"
            Case 1227
                stTitle = "ERROR_ADDRESS_ALREADY_ASSOCIATED"
            Case 1228
                stTitle = "ERROR_ADDRESS_NOT_ASSOCIATED"
            Case 1229
                stTitle = "ERROR_CONNECTION_INVALID"
            Case 1230
                stTitle = "ERROR_CONNECTION_ACTIVE"
            Case 1231
                stTitle = "ERROR_NETWORK_UNREACHABLE"
            Case 1232
                stTitle = "ERROR_HOST_UNREACHABLE"
            Case 1233
                stTitle = "ERROR_PROTOCOL_UNREACHABLE"
            Case 1234
                stTitle = "ERROR_PORT_UNREACHABLE"
            Case 1235
                stTitle = "ERROR_REQUEST_ABORTED"
            Case 1236
                stTitle = "ERROR_CONNECTION_ABORTED"
            Case 1237
                stTitle = "ERROR_RETRY"
            Case 1238
                stTitle = "ERROR_CONNECTION_COUNT_LIMIT"
            Case 1239
                stTitle = "ERROR_LOGIN_TIME_RESTRICTION"
            Case 1240
                stTitle = "ERROR_LOGIN_WKSTA_RESTRICTION"
            Case 1241
                stTitle = "ERROR_INCORRECT_ADDRESS"
            Case 1242
                stTitle = "ERROR_ALREADY_REGISTERED"
            Case 1243
                stTitle = "ERROR_SERVICE_NOT_FOUND"
            Case 1244
                stTitle = "ERROR_NOT_AUTHENTICATED"
            Case 1245
                stTitle = "ERROR_NOT_LOGGED_ON"
            Case 1246
                stTitle = "ERROR_CONTINUE"
            Case 1247
                stTitle = "ERROR_ALREADY_INITIALIZED"
            Case 1248
                stTitle = "ERROR_NO_MORE_DEVICES"
            Case 1249
                stTitle = "ERROR_NO_SUCH_SITE"
            Case 1250
                stTitle = "ERROR_DOMAIN_CONTROLLER_EXISTS"
            Case 1251
                stTitle = "ERROR_DS_NOT_INSTALLED"
            Case 1252 To 1299
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1300
                stTitle = "ERROR_NOT_ALL_ASSIGNED"
            Case 1301
                stTitle = "ERROR_SOME_NOT_MAPPED"
            Case 1302
                stTitle = "ERROR_NO_QUOTAS_FOR_ACCOUNT"
            Case 1303
                stTitle = "ERROR_LOCAL_USER_SESSION_KEY"
            Case 1304
                stTitle = "ERROR_NULL_LM_PASSWORD"
            Case 1305
                stTitle = "ERROR_UNKNOWN_REVISION"
            Case 1306
                stTitle = "ERROR_REVISION_MISMATCH"
            Case 1307
                stTitle = "ERROR_INVALID_OWNER"
            Case 1308
                stTitle = "ERROR_INVALID_PRIMARY_GROUP"
            Case 1309
                stTitle = "ERROR_NO_IMPERSONATION_TOKEN"
            Case 1310
                stTitle = "ERROR_CANT_DISABLE_MANDATORY"
            Case 1311
                stTitle = "ERROR_NO_LOGON_SERVERS"
            Case 1312
                stTitle = "ERROR_NO_SUCH_LOGON_SESSION"
            Case 1313
                stTitle = "ERROR_NO_SUCH_PRIVILEGE"
            Case 1314
                stTitle = "ERROR_PRIVILEGE_NOT_HELD"
            Case 1315
                stTitle = "ERROR_INVALID_ACCOUNT_NAME"
            Case 1316
                stTitle = "ERROR_USER_EXISTS"
            Case 1317
                stTitle = "ERROR_NO_SUCH_USER"
            Case 1318
                stTitle = "ERROR_GROUP_EXISTS"
            Case 1319
                stTitle = "ERROR_NO_SUCH_GROUP"
            Case 1320
                stTitle = "ERROR_MEMBER_IN_GROUP"
            Case 1321
                stTitle = "ERROR_MEMBER_NOT_IN_GROUP"
            Case 1322
                stTitle = "ERROR_LAST_ADMIN"
            Case 1323
                stTitle = "ERROR_WRONG_PASSWORD"
            Case 1324
                stTitle = "ERROR_ILL_FORMED_PASSWORD"
            Case 1325
                stTitle = "ERROR_PASSWORD_RESTRICTION"
            Case 1326
                stTitle = "ERROR_LOGON_FAILURE"
            Case 1327
                stTitle = "ERROR_ACCOUNT_RESTRICTION"
            Case 1328
                stTitle = "ERROR_INVALID_LOGON_HOURS"
            Case 1329
                stTitle = "ERROR_INVALID_WORKSTATION"
            Case 1330
                stTitle = "ERROR_PASSWORD_EXPIRED"
            Case 1331
                stTitle = "ERROR_ACCOUNT_DISABLED"
            Case 1332
                stTitle = "ERROR_NONE_MAPPED"
            Case 1333
                stTitle = "ERROR_TOO_MANY_LUIDS_REQUESTED"
            Case 1334
                stTitle = "ERROR_LUIDS_EXHAUSTED"
            Case 1335
                stTitle = "ERROR_INVALID_SUB_AUTHORITY"
            Case 1336
                stTitle = "ERROR_INVALID_ACL"
            Case 1337
                stTitle = "ERROR_INVALID_SID"
            Case 1338
                stTitle = "ERROR_INVALID_SECURITY_DESCR"
            Case 1339
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1340
                stTitle = "ERROR_BAD_INHERITANCE_ACL"
            Case 1341
                stTitle = "ERROR_SERVER_DISABLED"
            Case 1342
                stTitle = "ERROR_SERVER_NOT_DISABLED"
            Case 1343
                stTitle = "ERROR_INVALID_ID_AUTHORITY"
            Case 1344
                stTitle = "ERROR_ALLOTTED_SPACE_EXCEEDED"
            Case 1345
                stTitle = "ERROR_INVALID_GROUP_ATTRIBUTES"
            Case 1346
                stTitle = "ERROR_BAD_IMPERSONATION_LEVEL"
            Case 1347
                stTitle = "ERROR_CANT_OPEN_ANONYMOUS"
            Case 1348
                stTitle = "ERROR_BAD_VALIDATION_CLASS"
            Case 1349
                stTitle = "ERROR_BAD_TOKEN_TYPE"
            Case 1350
                stTitle = "ERROR_NO_SECURITY_ON_OBJECT"
            Case 1351
                stTitle = "ERROR_CANT_ACCESS_DOMAIN_INFO"
            Case 1352
                stTitle = "ERROR_INVALID_SERVER_STATE"
            Case 1353
                stTitle = "ERROR_INVALID_DOMAIN_STATE"
            Case 1354
                stTitle = "ERROR_INVALID_DOMAIN_ROLE"
            Case 1355
                stTitle = "ERROR_NO_SUCH_DOMAIN"
            Case 1356
                stTitle = "ERROR_DOMAIN_EXISTS"
            Case 1357
                stTitle = "ERROR_DOMAIN_LIMIT_EXCEEDED"
            Case 1358
                stTitle = "ERROR_INTERNAL_DB_CORRUPTION"
            Case 1359
                stTitle = "ERROR_INTERNAL_ERROR"
            Case 1360
                stTitle = "ERROR_GENERIC_NOT_MAPPED"
            Case 1361
                stTitle = "ERROR_BAD_DESCRIPTOR_FORMAT"
            Case 1362
                stTitle = "ERROR_NOT_LOGON_PROCESS"
            Case 1363
                stTitle = "ERROR_LOGON_SESSION_EXISTS"
            Case 1364
                stTitle = "ERROR_NO_SUCH_PACKAGE"
            Case 1365
                stTitle = "ERROR_BAD_LOGON_SESSION_STATE"
            Case 1366
                stTitle = "ERROR_LOGON_SESSION_COLLISION"
            Case 1367
                stTitle = "ERROR_INVALID_LOGON_TYPE"
            Case 1368
                stTitle = "ERROR_CANNOT_IMPERSONATE"
            Case 1369
                stTitle = "ERROR_RXACT_INVALID_STATE"
            Case 1370
                stTitle = "ERROR_RXACT_COMMIT_FAILURE"
            Case 1371
                stTitle = "ERROR_SPECIAL_ACCOUNT"
            Case 1372
                stTitle = "ERROR_SPECIAL_GROUP"
            Case 1373
                stTitle = "ERROR_SPECIAL_USER"
            Case 1374
                stTitle = "ERROR_MEMBERS_PRIMARY_GROUP"
            Case 1375
                stTitle = "ERROR_TOKEN_ALREADY_IN_USE"
            Case 1376
                stTitle = "ERROR_NO_SUCH_ALIAS"
            Case 1377
                stTitle = "ERROR_MEMBER_NOT_IN_ALIAS"
            Case 1378
                stTitle = "ERROR_MEMBER_IN_ALIAS"
            Case 1379
                stTitle = "ERROR_ALIAS_EXISTS"
            Case 1380
                stTitle = "ERROR_LOGON_NOT_GRANTED"
            Case 1381
                stTitle = "ERROR_TOO_MANY_SECRETS"
            Case 1382
                stTitle = "ERROR_SECRET_TOO_LONG"
            Case 1383
                stTitle = "ERROR_INTERNAL_DB_ERROR"
            Case 1384
                stTitle = "ERROR_TOO_MANY_CONTEXT_IDS"
            Case 1385
                stTitle = "ERROR_LOGON_TYPE_NOT_GRANTED"
            Case 1386
                stTitle = "ERROR_NT_CROSS_ENCRYPTION_REQUIRED"
            Case 1387
                stTitle = "ERROR_NO_SUCH_MEMBER"
            Case 1388
                stTitle = "ERROR_INVALID_MEMBER"
            Case 1389
                stTitle = "ERROR_TOO_MANY_SIDS"
            Case 1390
                stTitle = "ERROR_LM_CROSS_ENCRYPTION_REQUIRED"
            Case 1391
                stTitle = "ERROR_NO_INHERITANCE"
            Case 1392
                stTitle = "ERROR_FILE_CORRUPT"
            Case 1393
                stTitle = "ERROR_DISK_CORRUPT"
            Case 1394
                stTitle = "ERROR_NO_USER_SESSION_KEY"
            Case 1395
                stTitle = "ERROR_LICENSE_QUOTA_EXCEEDED"
            Case 1396 To 1399
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1400
                stTitle = "ERROR_INVALID_WINDOW_HANDLE"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        End If
    ElseIf inErrNum > 1400 Then
        '* Handle Range 1400+
        If inErrNum > 1400 And inErrNum <= 1800 Then
            '* Handle Range 1401 to 1800
            Select Case inErrNum
            Case 1401
                stTitle = "ERROR_INVALID_MENU_HANDLE"
            Case 1402
                stTitle = "ERROR_INVALID_CURSOR_HANDLE"
            Case 1403
                stTitle = "ERROR_INVALID_ACCEL_HANDLE"
            Case 1404
                stTitle = "ERROR_INVALID_HOOK_HANDLE"
            Case 1405
                stTitle = "ERROR_INVALID_DWP_HANDLE"
            Case 1406
                stTitle = "ERROR_TLW_WITH_WSCHILD"
            Case 1407
                stTitle = "ERROR_CANNOT_FIND_WND_CLASS"
            Case 1408
                stTitle = "ERROR_WINDOW_OF_OTHER_THREAD"
            Case 1409
                stTitle = "ERROR_HOTKEY_ALREADY_REGISTERED"
            Case 1410
                stTitle = "ERROR_CLASS_ALREADY_EXISTS"
            Case 1411
                stTitle = "ERROR_CLASS_DOES_NOT_EXIST"
            Case 1412
                stTitle = "ERROR_CLASS_HAS_WINDOWS"
            Case 1413
                stTitle = "ERROR_INVALID_INDEX"
            Case 1414
                stTitle = "ERROR_INVALID_ICON_HANDLE"
            Case 1415
                stTitle = "ERROR_PRIVATE_DIALOG_INDEX"
            Case 1416
                stTitle = "ERROR_LISTBOX_ID_NOT_FOUND"
            Case 1417
                stTitle = "ERROR_NO_WILDCARD_CHARACTERS"
            Case 1418
                stTitle = "ERROR_CLIPBOARD_NOT_OPEN"
            Case 1419
                stTitle = "ERROR_HOTKEY_NOT_REGISTERED"
            Case 1420
                stTitle = "ERROR_WINDOW_NOT_DIALOG"
            Case 1421
                stTitle = "ERROR_CONTROL_ID_NOT_FOUND"
            Case 1422
                stTitle = "ERROR_INVALID_COMBOBOX_MESSAGE"
            Case 1423
                stTitle = "ERROR_WINDOW_NOT_COMBOBOX"
            Case 1424
                stTitle = "ERROR_INVALID_EDIT_HEIGHT"
            Case 1425
                stTitle = "ERROR_DC_NOT_FOUND"
            Case 1426
                stTitle = "ERROR_INVALID_HOOK_FILTER"
            Case 1427
                stTitle = "ERROR_INVALID_FILTER_PROC"
            Case 1428
                stTitle = "ERROR_HOOK_NEEDS_HMOD"
            Case 1429
                stTitle = "ERROR_GLOBAL_ONLY_HOOK"
            Case 1430
                stTitle = "ERROR_JOURNAL_HOOK_SET"
            Case 1431
                stTitle = "ERROR_HOOK_NOT_INSTALLED"
            Case 1432
                stTitle = "ERROR_INVALID_LB_MESSAGE"
            Case 1433
                stTitle = "ERROR_SETCOUNT_ON_BAD_LB"
            Case 1434
                stTitle = "ERROR_LB_WITHOUT_TABSTOPS"
            Case 1435
                stTitle = "ERROR_DESTROY_OBJECT_OF_OTHER_THREAD"
            Case 1436
                stTitle = "ERROR_CHILD_WINDOW_MENU"
            Case 1437
                stTitle = "ERROR_NO_SYSTEM_MENU"
            Case 1438
                stTitle = "ERROR_INVALID_MSGBOX_STYLE"
            Case 1439
                stTitle = "ERROR_INVALID_SPI_VALUE"
            Case 1440
                stTitle = "ERROR_SCREEN_ALREADY_LOCKED"
            Case 1441
                stTitle = "ERROR_HWNDS_HAVE_DIFF_PARENT"
            Case 1442
                stTitle = "ERROR_NOT_CHILD_WINDOW"
            Case 1443
                stTitle = "ERROR_INVALID_GW_COMMAND"
            Case 1444
                stTitle = "ERROR_INVALID_THREAD_ID"
            Case 1445
                stTitle = "ERROR_NON_MDICHILD_WINDOW"
            Case 1446
                stTitle = "ERROR_POPUP_ALREADY_ACTIVE"
            Case 1447
                stTitle = "ERROR_NO_SCROLLBARS"
            Case 1448
                stTitle = "ERROR_INVALID_SCROLLBAR_RANGE"
            Case 1449
                stTitle = "ERROR_INVALID_SHOWWIN_COMMAND"
            Case 1450
                stTitle = "ERROR_NO_SYSTEM_RESOURCES"
            Case 1451
                stTitle = "ERROR_NONPAGED_SYSTEM_RESOURCES"
            Case 1452
                stTitle = "ERROR_PAGED_SYSTEM_RESOURCES"
            Case 1453
                stTitle = "ERROR_WORKING_SET_QUOTA"
            Case 1454
                stTitle = "ERROR_PAGEFILE_QUOTA"
            Case 1455
                stTitle = "ERROR_COMMITMENT_LIMIT"
            Case 1456
                stTitle = "ERROR_MENU_ITEM_NOT_FOUND"
            Case 1457
                stTitle = "ERROR_INVALID_KEYBOARD_HANDLE"
            Case 1458
                stTitle = "ERROR_HOOK_TYPE_NOT_ALLOWED"
            Case 1459
                stTitle = "ERROR_REQUIRES_INTERACTIVE_WINDOWSTATION"
            Case 1460
                stTitle = "ERROR_TIMEOUT"
            Case 1461
                stTitle = "ERROR_INVALID_MONITOR_HANDLE"
            Case 1462 To 1499
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1500
                stTitle = "ERROR_EVENTLOG_FILE_CORRUPT"
            Case 1501
                stTitle = "ERROR_EVENTLOG_CANT_START"
            Case 1502
                stTitle = "ERROR_LOG_FILE_FULL"
            Case 1503
                stTitle = "ERROR_EVENTLOG_FILE_CHANGED"
            Case 1504 To 1600
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1601
                stTitle = "ERROR_INSTALL_SERVICE"
            Case 1602
                stTitle = "ERROR_INSTALL_USEREXIT"
            Case 1603
                stTitle = "ERROR_INSTALL_FAILURE"
            Case 1604
                stTitle = "ERROR_INSTALL_SUSPEND"
            Case 1605
                stTitle = "ERROR_UNKNOWN_PRODUCT"
            Case 1606
                stTitle = "ERROR_UNKNOWN_FEATURE"
            Case 1607
                stTitle = "ERROR_UNKNOWN_COMPONENT"
            Case 1608
                stTitle = "ERROR_UNKNOWN_PROPERTY"
            Case 1609
                stTitle = "ERROR_INVALID_HANDLE_STATE"
            Case 1610
                stTitle = "ERROR_BAD_CONFIGURATION"
            Case 1611
                stTitle = "ERROR_INDEX_ABSENT"
            Case 1612
                stTitle = "ERROR_INSTALL_SOURCE_ABSENT"
            Case 1613
                stTitle = "ERROR_BAD_DATABASE_VERSION"
            Case 1614
                stTitle = "ERROR_PRODUCT_UNINSTALLED"
            Case 1615
                stTitle = "ERROR_BAD_QUERY_SYNTAX"
            Case 1616
                stTitle = "ERROR_INVALID_FIELD"
            Case 1617 To 1699
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1700
                stTitle = "RPC_S_INVALID_STRING_BINDING"
            Case 1701
                stTitle = "RPC_S_WRONG_KIND_OF_BINDING"
            Case 1702
                stTitle = "RPC_S_INVALID_BINDING"
            Case 1703
                stTitle = "RPC_S_PROTSEQ_NOT_SUPPORTED"
            Case 1704
                stTitle = "RPC_S_INVALID_RPC_PROTSEQ"
            Case 1705
                stTitle = "RPC_S_INVALID_STRING_UUID"
            Case 1706
                stTitle = "RPC_S_INVALID_ENDPOINT_FORMAT"
            Case 1707
                stTitle = "RPC_S_INVALID_NET_ADDR"
            Case 1708
                stTitle = "RPC_S_NO_ENDPOINT_FOUND"
            Case 1709
                stTitle = "RPC_S_INVALID_TIMEOUT"
            Case 1710
                stTitle = "RPC_S_OBJECT_NOT_FOUND"
            Case 1711
                stTitle = "RPC_S_ALREADY_REGISTERED"
            Case 1712
                stTitle = "RPC_S_TYPE_ALREADY_REGISTERED"
            Case 1713
                stTitle = "RPC_S_ALREADY_LISTENING"
            Case 1714
                stTitle = "RPC_S_NO_PROTSEQS_REGISTERED"
            Case 1715
                stTitle = "RPC_S_NOT_LISTENING"
            Case 1716
                stTitle = "RPC_S_UNKNOWN_MGR_TYPE"
            Case 1717
                stTitle = "RPC_S_UNKNOWN_IF"
            Case 1718
                stTitle = "RPC_S_NO_BINDINGS"
            Case 1719
                stTitle = "RPC_S_NO_PROTSEQS"
            Case 1720
                stTitle = "RPC_S_CANT_CREATE_ENDPOINT"
            Case 1721
                stTitle = "RPC_S_OUT_OF_RESOURCES"
            Case 1722
                stTitle = "RPC_S_SERVER_UNAVAILABLE"
            Case 1723
                stTitle = "RPC_S_SERVER_TOO_BUSY"
            Case 1724
                stTitle = "RPC_S_INVALID_NETWORK_OPTIONS"
            Case 1725
                stTitle = "RPC_S_NO_CALL_ACTIVE"
            Case 1726
                stTitle = "RPC_S_CALL_FAILED"
            Case 1727
                stTitle = "RPC_S_CALL_FAILED_DNE"
            Case 1728
                stTitle = "RPC_S_PROTOCOL_ERROR"
            Case 1729
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1730
                stTitle = "RPC_S_UNSUPPORTED_TRANS_SYN"
            Case 1731
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1732
                stTitle = "RPC_S_UNSUPPORTED_TYPE"
            Case 1733
                stTitle = "RPC_S_INVALID_TAG"
            Case 1734
                stTitle = "RPC_S_INVALID_BOUND"
            Case 1735
                stTitle = "RPC_S_NO_ENTRY_NAME"
            Case 1736
                stTitle = "RPC_S_INVALID_NAME_SYNTAX"
            Case 1737
                stTitle = "RPC_S_UNSUPPORTED_NAME_SYNTAX"
            Case 1738
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1739
                stTitle = "RPC_S_UUID_NO_ADDRESS"
            Case 1740
                stTitle = "RPC_S_DUPLICATE_ENDPOINT"
            Case 1741
                stTitle = "RPC_S_UNKNOWN_AUTHN_TYPE"
            Case 1742
                stTitle = "RPC_S_MAX_CALLS_TOO_SMALL"
            Case 1743
                stTitle = "RPC_S_STRING_TOO_LONG"
            Case 1744
                stTitle = "RPC_S_PROTSEQ_NOT_FOUND"
            Case 1745
                stTitle = "RPC_S_PROCNUM_OUT_OF_RANGE"
            Case 1746
                stTitle = "RPC_S_BINDING_HAS_NO_AUTH"
            Case 1747
                stTitle = "RPC_S_UNKNOWN_AUTHN_SERVICE"
            Case 1748
                stTitle = "RPC_S_UNKNOWN_AUTHN_LEVEL"
            Case 1749
                stTitle = "RPC_S_INVALID_AUTH_IDENTITY"
            Case 1750
                stTitle = "RPC_S_UNKNOWN_AUTHZ_SERVICE"
            Case 1751
                stTitle = "EPT_S_INVALID_ENTRY"
            Case 1752
                stTitle = "EPT_S_CANT_PERFORM_OP"
            Case 1753
                stTitle = "EPT_S_NOT_REGISTERED"
            Case 1754
                stTitle = "RPC_S_NOTHING_TO_EXPORT"
            Case 1755
                stTitle = "RPC_S_INCOMPLETE_NAME"
            Case 1756
                stTitle = "RPC_S_INVALID_VERS_OPTION"
            Case 1757
                stTitle = "RPC_S_NO_MORE_MEMBERS"
            Case 1758
                stTitle = "RPC_S_NOT_ALL_OBJS_UNEXPORTED"
            Case 1759
                stTitle = "RPC_S_INTERFACE_NOT_FOUND"
            Case 1760
                stTitle = "RPC_S_ENTRY_ALREADY_EXISTS"
            Case 1761
                stTitle = "RPC_S_ENTRY_NOT_FOUND"
            Case 1762
                stTitle = "RPC_S_NAME_SERVICE_UNAVAILABLE"
            Case 1763
                stTitle = "RPC_S_INVALID_NAF_ID"
            Case 1764
                stTitle = "RPC_S_CANNOT_SUPPORT"
            Case 1765
                stTitle = "RPC_S_NO_CONTEXT_AVAILABLE"
            Case 1766
                stTitle = "RPC_S_INTERNAL_ERROR"
            Case 1767
                stTitle = "RPC_S_ZERO_DIVIDE"
            Case 1768
                stTitle = "RPC_S_ADDRESS_ERROR"
            Case 1769
                stTitle = "RPC_S_FP_DIV_ZERO"
            Case 1770
                stTitle = "RPC_S_FP_UNDERFLOW"
            Case 1771
                stTitle = "RPC_S_FP_OVERFLOW"
            Case 1772
                stTitle = "RPC_X_NO_MORE_ENTRIES"
            Case 1773
                stTitle = "RPC_X_SS_CHAR_TRANS_OPEN_FAIL"
            Case 1774
                stTitle = "RPC_X_SS_CHAR_TRANS_SHORT_FILE"
            Case 1775
                stTitle = "RPC_X_SS_IN_NULL_CONTEXT"
            Case 1776
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1777
                stTitle = "RPC_X_SS_CONTEXT_DAMAGED"
            Case 1778
                stTitle = "RPC_X_SS_HANDLES_MISMATCH"
            Case 1779
                stTitle = "RPC_X_SS_CANNOT_GET_CALL_HANDLE"
            Case 1780
                stTitle = "RPC_X_NULL_REF_POINTER"
            Case 1781
                stTitle = "RPC_X_ENUM_VALUE_OUT_OF_RANGE"
            Case 1782
                stTitle = "RPC_X_BYTE_COUNT_TOO_SMALL"
            Case 1783
                stTitle = "RPC_X_BAD_STUB_DATA"
            Case 1784
                stTitle = "ERROR_INVALID_USER_BUFFER"
            Case 1785
                stTitle = "ERROR_UNRECOGNIZED_MEDIA"
            Case 1786
                stTitle = "ERROR_NO_TRUST_LSA_SECRET"
            Case 1787
                stTitle = "ERROR_NO_TRUST_SAM_ACCOUNT"
            Case 1788
                stTitle = "ERROR_TRUSTED_DOMAIN_FAILURE"
            Case 1789
                stTitle = "ERROR_TRUSTED_RELATIONSHIP_FAILURE"
            Case 1790
                stTitle = "ERROR_TRUST_FAILURE"
            Case 1791
                stTitle = "RPC_S_CALL_IN_PROGRESS"
            Case 1792
                stTitle = "ERROR_NETLOGON_NOT_STARTED"
            Case 1793
                stTitle = "ERROR_ACCOUNT_EXPIRED"
            Case 1794
                stTitle = "ERROR_REDIRECTOR_HAS_OPEN_HANDLES"
            Case 1795
                stTitle = "ERROR_PRINTER_DRIVER_ALREADY_INSTALLED"
            Case 1796
                stTitle = "ERROR_UNKNOWN_PORT"
            Case 1797
                stTitle = "ERROR_UNKNOWN_PRINTER_DRIVER"
            Case 1798
                stTitle = "ERROR_UNKNOWN_PRINTPROCESSOR"
            Case 1799
                stTitle = "ERROR_INVALID_SEPARATOR_FILE"
            Case 1800
                stTitle = "ERROR_INVALID_PRIORITY"
            End Select
        ElseIf inErrNum > 1800 And inErrNum <= 1999 Then
            '* Handle Range 1801 to 1999
            Select Case inErrNum
            Case 1801
                stTitle = "ERROR_INVALID_PRINTER_NAME"
            Case 1802
                stTitle = "ERROR_PRINTER_ALREADY_EXISTS"
            Case 1803
                stTitle = "ERROR_INVALID_PRINTER_COMMAND"
            Case 1804
                stTitle = "ERROR_INVALID_DATATYPE"
            Case 1805
                stTitle = "ERROR_INVALID_ENVIRONMENT"
            Case 1806
                stTitle = "RPC_S_NO_MORE_BINDINGS"
            Case 1807
                stTitle = "ERROR_NOLOGON_INTERDOMAIN_TRUST_ACCOUNT"
            Case 1808
                stTitle = "ERROR_NOLOGON_WORKSTATION_TRUST_ACCOUNT"
            Case 1809
                stTitle = "ERROR_NOLOGON_SERVER_TRUST_ACCOUNT"
            Case 1810
                stTitle = "ERROR_DOMAIN_TRUST_INCONSISTENT"
            Case 1811
                stTitle = "ERROR_SERVER_HAS_OPEN_HANDLES"
            Case 1812
                stTitle = "ERROR_RESOURCE_DATA_NOT_FOUND"
            Case 1813
                stTitle = "ERROR_RESOURCE_TYPE_NOT_FOUND"
            Case 1814
                stTitle = "ERROR_RESOURCE_NAME_NOT_FOUND"
            Case 1815
                stTitle = "ERROR_RESOURCE_LANG_NOT_FOUND"
            Case 1816
                stTitle = "ERROR_NOT_ENOUGH_QUOTA"
            Case 1817
                stTitle = "RPC_S_NO_INTERFACES"
            Case 1818
                stTitle = "RPC_S_CALL_CANCELLED"
            Case 1819
                stTitle = "RPC_S_BINDING_INCOMPLETE"
            Case 1820
                stTitle = "RPC_S_COMM_FAILURE"
            Case 1821
                stTitle = "RPC_S_UNSUPPORTED_AUTHN_LEVEL"
            Case 1822
                stTitle = "RPC_S_NO_PRINC_NAME"
            Case 1823
                stTitle = "RPC_S_NOT_RPC_ERROR"
            Case 1824
                stTitle = "RPC_S_UUID_LOCAL_ONLY"
            Case 1825
                stTitle = "RPC_S_SEC_PKG_ERROR"
            Case 1826
                stTitle = "RPC_S_NOT_CANCELLED"
            Case 1827
                stTitle = "RPC_X_INVALID_ES_ACTION"
            Case 1828
                stTitle = "RPC_X_WRONG_ES_VERSION"
            Case 1829
                stTitle = "RPC_X_WRONG_STUB_VERSION"
            Case 1830
                stTitle = "RPC_X_INVALID_PIPE_OBJECT"
            Case 1831
                stTitle = "RPC_X_WRONG_PIPE_ORDER"
            Case 1832
                stTitle = "RPC_X_WRONG_PIPE_VERSION"
            Case 1833 To 1897
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 1898
                stTitle = "RPC_S_GROUP_MEMBER_NOT_FOUND"
            Case 1899
                stTitle = "EPT_S_CANT_CREATE"
            Case 1900
                stTitle = "RPC_S_INVALID_OBJECT"
            Case 1901
                stTitle = "ERROR_INVALID_TIME"
            Case 1902
                stTitle = "ERROR_INVALID_FORM_NAME"
            Case 1903
                stTitle = "ERROR_INVALID_FORM_SIZE"
            Case 1904
                stTitle = "ERROR_ALREADY_WAITING"
            Case 1905
                stTitle = "ERROR_PRINTER_DELETED"
            Case 1906
                stTitle = "ERROR_INVALID_PRINTER_STATE"
            Case 1907
                stTitle = "ERROR_PASSWORD_MUST_CHANGE"
            Case 1908
                stTitle = "ERROR_DOMAIN_CONTROLLER_NOT_FOUND"
            Case 1909
                stTitle = "ERROR_ACCOUNT_LOCKED_OUT"
            Case 1910
                stTitle = "OR_INVALID_OXID"
            Case 1911
                stTitle = "OR_INVALID_OID"
            Case 1912
                stTitle = "OR_INVALID_SET"
            Case 1913
                stTitle = "RPC_S_SEND_INCOMPLETE"
            Case 1914
                stTitle = "RPC_S_INVALID_ASYNC_HANDLE"
            Case 1915
                stTitle = "RPC_S_INVALID_ASYNC_CALL"
            Case 1916
                stTitle = "RPC_X_PIPE_CLOSED"
            Case 1917
                stTitle = "RPC_X_PIPE_DISCIPLINE_ERROR"
            Case 1918
                stTitle = "RPC_X_PIPE_EMPTY"
            Case 1919
                stTitle = "ERROR_NO_SITENAME"
            Case 1920
                stTitle = "ERROR_CANT_ACCESS_FILE"
            Case 1921
                stTitle = "ERROR_CANT_RESOLVE_FILENAME"
            Case 1922
                stTitle = "ERROR_DS_MEMBERSHIP_EVALUATED_LOCALLY"
            Case 1923
                stTitle = "ERROR_DS_NO_ATTRIBUTE_OR_VALUE"
            Case 1924
                stTitle = "ERROR_DS_INVALID_ATTRIBUTE_SYNTAX"
            Case 1925
                stTitle = "ERROR_DS_ATTRIBUTE_TYPE_UNDEFINED"
            Case 1926
                stTitle = "ERROR_DS_ATTRIBUTE_OR_VALUE_EXISTS"
            Case 1927
                stTitle = "ERROR_DS_BUSY"
            Case 1928
                stTitle = "ERROR_DS_UNAVAILABLE"
            Case 1929
                stTitle = "ERROR_DS_NO_RIDS_ALLOCATED"
            Case 1930
                stTitle = "ERROR_DS_NO_MORE_RIDS"
            Case 1931
                stTitle = "ERROR_DS_INCORRECT_ROLE_OWNER"
            Case 1932
                stTitle = "ERROR_DS_RIDMGR_INIT_ERROR"
            Case 1933
                stTitle = "ERROR_DS_OBJ_CLASS_VIOLATION"
            Case 1934
                stTitle = "ERROR_DS_CANT_ON_NON_LEAF"
            Case 1935
                stTitle = "ERROR_DS_CANT_ON_RDN"
            Case 1936
                stTitle = "ERROR_DS_CANT_MOD_OBJ_CLASS"
            Case 1937
                stTitle = "ERROR_DS_CROSS_DOM_MOVE_ERROR"
            Case 1938
                stTitle = "ERROR_DS_GC_NOT_AVAILABLE"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 1999 And inErrNum <= 2999 Then
            '* Handle Range 2000 to 2999
            Select Case inErrNum
            Case 2000
                stTitle = "ERROR_INVALID_PIXEL_FORMAT"
            Case 2001
                stTitle = "ERROR_BAD_DRIVER"
            Case 2002
                stTitle = "ERROR_INVALID_WINDOW_STYLE"
            Case 2003
                stTitle = "ERROR_METAFILE_NOT_SUPPORTED"
            Case 2004
                stTitle = "ERROR_TRANSFORM_NOT_SUPPORTED"
            Case 2005
                stTitle = "ERROR_CLIPPING_NOT_SUPPORTED"
            Case 2006 To 2107
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 2108
                stTitle = "ERROR_CONNECTED_OTHER_PASSWORD"
            Case 2109 To 2201
                stTitle = "Error Code " & CStr(inErrNum) & "Not Listed."
            Case 2202
                stTitle = "ERROR_BAD_USERNAME"
            Case 2203 To 2249
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 2250
                stTitle = "ERROR_NOT_CONNECTED"
            Case 2251 To 2299
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 2300
                stTitle = "ERROR_INVALID_CMM"
            Case 2301
                stTitle = "ERROR_INVALID_PROFILE"
            Case 2302
                stTitle = "ERROR_TAG_NOT_FOUND"
            Case 2303
                stTitle = "ERROR_TAG_NOT_PRESENT"
            Case 2304
                stTitle = "ERROR_DUPLICATE_TAG"
            Case 2305
                stTitle = "ERROR_PROFILE_NOT_ASSOCIATED_WITH_DEVICE"
            Case 2306
                stTitle = "ERROR_PROFILE_NOT_FOUND"
            Case 2307
                stTitle = "ERROR_INVALID_COLORSPACE"
            Case 2308
                stTitle = "ERROR_ICM_NOT_ENABLED"
            Case 2309
                stTitle = "ERROR_DELETING_ICM_XFORM"
            Case 2310
                stTitle = "ERROR_INVALID_TRANSFORM"
            Case 2311 To 2400
                stTitle = "Error Code" & CStr(inErrNum) & " Not Listed."
            Case 2401
                stTitle = "ERROR_OPEN_FILES"
            Case 2402
                stTitle = "ERROR_ACTIVE_CONNECTIONS"
            Case 2403
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 2404
                stTitle = "ERROR_DEVICE_IN_USE"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 2999 And inErrNum <= 3999 Then
            '* Handle Range 3000 to 3999
            Select Case inErrNum
            Case 3000
                stTitle = "ERROR_UNKNOWN_PRINT_MONITOR"
            Case 3001
                stTitle = "ERROR_PRINTER_DRIVER_IN_USE"
            Case 3002
                stTitle = "ERROR_SPOOL_FILE_NOT_FOUND"
            Case 3003
                stTitle = "ERROR_SPL_NO_STARTDOC"
            Case 3004
                stTitle = "ERROR_SPL_NO_ADDJOB"
            Case 3005
                stTitle = "ERROR_PRINT_PROCESSOR_ALREADY_INSTALLED"
            Case 3006
                stTitle = "ERROR_PRINT_MONITOR_ALREADY_INSTALLED"
            Case 3007
                stTitle = "ERROR_INVALID_PRINT_MONITOR"
            Case 3008
                stTitle = "ERROR_PRINT_MONITOR_IN_USE"
            Case 3009
                stTitle = "ERROR_PRINTER_HAS_JOBS_QUEUED"
            Case 3010
                stTitle = "ERROR_SUCCESS_REBOOT_REQUIRED"
            Case 3011
                stTitle = "ERROR_SUCCESS_RESTART_REQUIRED"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 3999 And inErrNum <= 4999 Then
            '* Handle Range 4000 to 4999
            Select Case inErrNum
            Case 4000
                stTitle = "ERROR_WINS_INTERNAL"
            Case 4001
                stTitle = "ERROR_CAN_NOT_DEL_LOCAL_WINS"
            Case 4002
                stTitle = "ERROR_STATIC_INIT = 4002"
            Case 4003
                stTitle = "ERROR_INC_BACKUP"
            Case 4004
                stTitle = "ERROR_FULL_BACKUP"
            Case 4005
                stTitle = "ERROR_REC_NON_EXISTENT"
            Case 4006
                stTitle = "ERROR_RPL_NOT_ALLOWED"
            Case 4007 To 4099
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 4100
                stTitle = "ERROR_DHCP_ADDRESS_CONFLICT"
            Case 4101 To 4199
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 4200
                stTitle = "ERROR_WMI_GUID_NOT_FOUND"
            Case 4201
                stTitle = "ERROR_WMI_INSTANCE_NOT_FOUND"
            Case 4202
                stTitle = "ERROR_WMI_ITEMID_NOT_FOUND"
            Case 4203
                stTitle = "ERROR_WMI_TRY_AGAIN"
            Case 4204
                stTitle = "ERROR_WMI_DP_NOT_FOUND"
            Case 4205
                stTitle = "ERROR_WMI_UNRESOLVED_INSTANCE_REF"
            Case 4206
                stTitle = "ERROR_WMI_ALREADY_ENABLED"
            Case 4207
                stTitle = "ERROR_WMI_GUID_DISCONNECTED"
            Case 4208
                stTitle = "ERROR_WMI_SERVER_UNAVAILABLE"
            Case 4209
                stTitle = "ERROR_WMI_DP_FAILED"
            Case 4210
                stTitle = "ERROR_WMI_INVALID_MOF"
            Case 4211
                stTitle = "ERROR_WMI_INVALID_REGINFO"
            Case 4212 To 4299
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 4300
                stTitle = "ERROR_INVALID_MEDIA"
            Case 4301
                stTitle = "ERROR_INVALID_LIBRARY"
            Case 4302
                stTitle = "ERROR_INVALID_MEDIA_POOL"
            Case 4303
                stTitle = "ERROR_DRIVE_MEDIA_MISMATCH"
            Case 4304
                stTitle = "ERROR_MEDIA_OFFLINE"
            Case 4305
                stTitle = "ERROR_LIBRARY_OFFLINE"
            Case 4306
                stTitle = "ERROR_EMPTY"
            Case 4307
                stTitle = "ERROR_NOT_EMPTY"
            Case 4308
                stTitle = "ERROR_MEDIA_UNAVAILABLE"
            Case 4309
                stTitle = "ERROR_RESOURCE_DISABLED"
            Case 4310
                stTitle = "ERROR_INVALID_CLEANER"
            Case 4311
                stTitle = "ERROR_UNABLE_TO_CLEAN"
            Case 4312
                stTitle = "ERROR_OBJECT_NOT_FOUND"
            Case 4313
                stTitle = "ERROR_DATABASE_FAILURE"
            Case 4314
                stTitle = "ERROR_DATABASE_FULL"
            Case 4315
                stTitle = "ERROR_MEDIA_INCOMPATIBLE"
            Case 4316
                stTitle = "ERROR_RESOURCE_NOT_PRESENT"
            Case 4317
                stTitle = "ERROR_INVALID_OPERATION"
            Case 4318
                stTitle = "ERROR_MEDIA_NOT_AVAILABLE"
            Case 4319
                stTitle = "ERROR_DEVICE_NOT_AVAILABLE"
            Case 4320
                stTitle = "ERROR_REQUEST_REFUSED"
            Case 4321 To 4349
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 4350
                stTitle = "ERROR_FILE_OFFLINE"
            Case 4351
                stTitle = "ERROR_REMOTE_STORAGE_NOT_ACTIVE"
            Case 4352
                stTitle = "ERROR_REMOTE_STORAGE_MEDIA_ERROR"
            Case 4353 To 4389
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 4390
                stTitle = "ERROR_NOT_A_REPARSE_POINT"
            Case 4391
                stTitle = "ERROR_REPARSE_ATTRIBUTE_CONFLICT"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 4999 And inErrNum <= 5999 Then
            '* Handle Range 5000 to 5999
            Select Case inErrNum
            Case 5000
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 5001
                stTitle = "ERROR_DEPENDENT_RESOURCE_EXISTS"
            Case 5002
                stTitle = "ERROR_DEPENDENCY_NOT_FOUND"
            Case 5003
                stTitle = "ERROR_DEPENDENCY_ALREADY_EXISTS"
            Case 5004
                stTitle = "ERROR_RESOURCE_NOT_ONLINE"
            Case 5005
                stTitle = "ERROR_HOST_NODE_NOT_AVAILABLE"
            Case 5006
                stTitle = "ERROR_RESOURCE_NOT_AVAILABLE"
            Case 5007
                stTitle = "ERROR_RESOURCE_NOT_FOUND"
            Case 5008
                stTitle = "ERROR_SHUTDOWN_CLUSTER"
            Case 5009
                stTitle = "ERROR_CANT_EVICT_ACTIVE_NODE"
            Case 5010
                stTitle = "ERROR_OBJECT_ALREADY_EXISTS"
            Case 5011
                stTitle = "ERROR_OBJECT_IN_LIST"
            Case 5012
                stTitle = "ERROR_GROUP_NOT_AVAILABLE"
            Case 5013
                stTitle = "ERROR_GROUP_NOT_FOUND"
            Case 5014
                stTitle = "ERROR_GROUP_NOT_ONLINE"
            Case 5015
                stTitle = "ERROR_HOST_NODE_NOT_RESOURCE_OWNER"
            Case 5016
                stTitle = "ERROR_HOST_NODE_NOT_GROUP_OWNER"
            Case 5017
                stTitle = "ERROR_RESMON_CREATE_FAILED"
            Case 5018
                stTitle = "ERROR_RESMON_ONLINE_FAILED"
            Case 5019
                stTitle = "ERROR_RESOURCE_ONLINE"
            Case 5020
                stTitle = "ERROR_QUORUM_RESOURCE"
            Case 5021
                stTitle = "ERROR_NOT_QUORUM_CAPABLE"
            Case 5022
                stTitle = "ERROR_CLUSTER_SHUTTING_DOWN"
            Case 5023
                stTitle = "ERROR_INVALID_STATE"
            Case 5024
                stTitle = "ERROR_RESOURCE_PROPERTIES_STORED"
            Case 5025
                stTitle = "ERROR_NOT_QUORUM_CLASS"
            Case 5026
                stTitle = "ERROR_CORE_RESOURCE"
            Case 5027
                stTitle = "ERROR_QUORUM_RESOURCE_ONLINE_FAILED"
            Case 5028
                stTitle = "ERROR_QUORUMLOG_OPEN_FAILED"
            Case 5029
                stTitle = "ERROR_CLUSTERLOG_CORRUPT"
            Case 5030
                stTitle = "ERROR_CLUSTERLOG_RECORD_EXCEEDS_MAXSIZE"
            Case 5031
                stTitle = "ERROR_CLUSTERLOG_EXCEEDS_MAXSIZE"
            Case 5032
                stTitle = "ERROR_CLUSTERLOG_CHKPOINT_NOT_FOUND"
            Case 5033
                stTitle = "ERROR_CLUSTERLOG_NOT_ENOUGH_SPACE"
            Case 5034 To 5999
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            End Select
        ElseIf inErrNum > 5999 And inErrNum <= 6999 Then
            '* Handle Range 6000 to 6999
            Select Case inErrNum
            Case 6000
                stTitle = "ERROR_ENCRYPTION_FAILED"
            Case 6001
                stTitle = "ERROR_DECRYPTION_FAILED"
            Case 6002
                stTitle = "ERROR_FILE_ENCRYPTED"
            Case 6003
                stTitle = "ERROR_NO_RECOVERY_POLICY"
            Case 6004
                stTitle = "ERROR_NO_EFS"
            Case 6005
                stTitle = "ERROR_WRONG_EFS"
            Case 6006
                stTitle = "ERROR_NO_USER_KEYS"
            Case 6007
                stTitle = "ERROR_FILE_NOT_ENCRYPTED"
            Case 6008
                stTitle = "ERROR_NOT_EXPORT_FORMAT"
            Case 6009 To 6117
                stTitle = "Error Code " & CStr(inErrNum) & " Not Listed."
            Case 6118
                stTitle = "ERROR_NO_BROWSER_SERVERS_FOUND"
            Case Else
                stTitle = "Error Code " & CStr(inErrNum) & " is out of range."
            End Select
        Else
            stTitle = "Error Number Out Of Range (Greater Than 6999)"
        End If
    End If
    
    GetErrorTitle = stTitle
    
End Function
