#[macro_use]
extern crate lazy_static;
use chrono::prelude::*;
use std::time::Instant;
use uuid::Uuid;
use warp::{Filter, Rejection, Reply};

mod excel;

type Result<T> = std::result::Result<T, Rejection>;

lazy_static! {
    static ref THINGS: Vec<Thing> = create_things();
}

#[derive(Clone, Debug)]
pub struct Thing {
    pub id: String,
    pub start_date: DateTime<Utc>,
    pub end_date: DateTime<Utc>,
    pub project: String,
    pub name: String,
    pub text: String,
}

#[tokio::main]
async fn main() {
    let report_route = warp::path("report")
        .and(warp::get())
        .and_then(report_handler);

    println!("Server started at localhost:8080");
    warp::serve(report_route).run(([0, 0, 0, 0], 8080)).await;
}

async fn report_handler() -> Result<impl Reply> {
    let now = Instant::now();
    let result = tokio::task::spawn_blocking(move || excel::create_xlsx(THINGS.to_vec()))
        .await
        .expect("can create result");
    println!("report took: {:?}", now.elapsed());
    Ok(result)
}

fn create_things() -> Vec<Thing> {
    let mut result: Vec<Thing> = vec![];
    for _ in 0..1000 {
        result.push(Thing {
            id: random_string(),
            start_date: Utc::now(),
            end_date: Utc::now(),
            project: random_string(),
            name: random_string(),
            text: random_string(),
        });
    }
    result
}

fn random_string() -> String {
    Uuid::new_v4().to_string()
}
