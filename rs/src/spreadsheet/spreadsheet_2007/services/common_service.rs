use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::spreadsheet_2007::services::{CalculationChainPart, ShareStringPart, StylePart};
use anyhow::{Context, Error as AnyError, Result as AnyResult};

#[derive(Debug)]
pub(crate) struct CommonServices {
    calculation_chain: CalculationChainPart,
    share_string: ShareStringPart,
    style: StylePart,
}

impl CommonServices {
    pub(crate) fn new(
        calculation_chain: CalculationChainPart,
        share_string: ShareStringPart,
        style: StylePart,
    ) -> Self {
        Self {
            calculation_chain,
            share_string,
            style,
        }
    }
    pub(crate) fn close_service(&mut self) -> AnyResult<(), AnyError> {
        self.calculation_chain
            .close_document()
            .context("Common Service Calculation Chain Close Failed")?;
        self.share_string
            .close_document()
            .context("Common Service Share String Close Failed")?;
        self.style
            .close_document()
            .context("Common Style Chain Close Failed")?;
        Ok(())
    }
}