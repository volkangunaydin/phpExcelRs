#![cfg_attr(windows, feature(abi_vectorcall))]

//! excel_rs — Excel reader/writer extension for PHP.
//!
//! Reading is backed by `calamine` (xlsx, xlsm, xlsb, xls, ods) and writing by
//! `rust_xlsxwriter` (xlsx).

mod reader;
mod writer;

use ext_php_rs::{exception::PhpException, prelude::*, zend::ce};

/// `ExcelRs\Exception` — thrown for every error raised by this extension.
#[php_class]
#[php(name = "ExcelRs\\Exception")]
#[php(extends(ce = ce::exception, stub = "\\Exception"))]
#[derive(Default)]
pub struct ExcelException;

/// Build an `ExcelRs\Exception` from anything printable.
pub(crate) fn err(message: impl std::fmt::Display) -> PhpException {
    PhpException::from_class::<ExcelException>(message.to_string())
}

#[php_module]
pub fn get_module(module: ModuleBuilder) -> ModuleBuilder {
    module
        .class::<ExcelException>()
        .class::<reader::Reader>()
        .class::<writer::Formula>()
        .class::<writer::Writer>()
}
