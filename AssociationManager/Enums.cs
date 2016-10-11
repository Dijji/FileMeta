// Copyright (c) 2015, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

namespace FileMetadataAssociationManager
{
    public enum ProfileControls
    {
        Unknown,
        Groups,
        Properties,
        FullDetails,
        PreviewDetails,
        InfoTip
    }

    public enum PropType
    {
        Group = 1,
        Normal,
    }

    public enum WindowsErrorCodes
    {
        ERROR_FILE_NOT_FOUND = 2,
        ERROR_ACCESS_DENIED = 5,
        ERROR_INVALID_PARAMETER = 87,
        ERROR_XML_PARSE_ERROR = 1465,
    }
}
