//go:build windows
// +build windows

package taskmaster

import (
	"errors"
	"fmt"
	"syscall"

	ole "github.com/go-ole/go-ole"
)

var (
	ErrTargetUnsupported    = errors.New("error connecting to the Task Scheduler service: cannot connect to the XP or server 2003 computer")
	ErrConnectionFailure    = errors.New("error connecting to the Task Scheduler service: cannot connect to target computer")
	ErrInvalidPath          = errors.New(`path must start with root folder "\"`)
	ErrNoActions            = errors.New("definition must have at least one action")
	ErrInvalidPrinciple     = errors.New("both UserId and GroupId are defined for the principal; they are mutually exclusive")
	ErrRunningTaskCompleted = errors.New("the running task completed while it was getting parsed")
)

func getTaskSchedulerError(err error) error {
	errCode, codeErr := getOLEErrorCode(err)
	if codeErr != nil {
		return fmt.Errorf("task scheduler error: %v (unable to get specific error code: %v)", err, codeErr)
	}

	switch errCode {
	case 50:
		return ErrTargetUnsupported
	case 0x80070032, 53:
		return ErrConnectionFailure
	default:
		return syscall.Errno(errCode)
	}
}

func getRunningTaskError(err error) error {
	errCode, codeErr := getOLEErrorCode(err)
	if codeErr != nil {
		return fmt.Errorf("running task error: %v (unable to get specific error code: %v)", err, codeErr)
	}

	if errCode == 0x8004130B {
		return ErrRunningTaskCompleted
	}
	return syscall.Errno(errCode)
}

func getOLEErrorCode(err error) (uint32, error) {
	if oleErr, ok := err.(*ole.OleError); ok {
		if oleErr.Code() != 0 {
			return uint32(oleErr.Code()), nil
		}

		if subErr := oleErr.SubError(); subErr != nil {
			if excepInfo, ok := subErr.(ole.EXCEPINFO); ok {
				return uint32(excepInfo.SCODE()), nil
			}
		}
	}

	return 0, errors.New("could not determine OLE error code")
}
