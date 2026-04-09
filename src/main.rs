mod error;
mod event;
mod outlook;

fn main() {
    env_logger::init();

    let now   = chrono::Utc::now();
    let start = now - chrono::Duration::days(outlook::DEFAULT_PAST_DAYS);
    let end   = now + chrono::Duration::days(outlook::DEFAULT_FUTURE_DAYS);

    match outlook::read_calendar_events(start, end) {
        Ok(events) => {
            println!(
                "\n{} calendar items ({} days ago → {} days ahead)\n",
                events.len(),
                outlook::DEFAULT_PAST_DAYS,
                outlook::DEFAULT_FUTURE_DAYS,
            );
            for event in &events {
                println!("{event}\n");
            }
        }
        Err(e) => {
            eprintln!("Error reading calendar: {e}");
            std::process::exit(1);
        }
    }
}
